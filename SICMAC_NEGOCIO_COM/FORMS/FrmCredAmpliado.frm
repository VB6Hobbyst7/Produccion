VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCredAmpliado 
   Caption         =   "Ampliación de Creditos"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6225
   Icon            =   "FrmCredAmpliado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6225
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   6105
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar..."
         Height          =   345
         Left            =   4440
         TabIndex        =   6
         Top             =   240
         Width           =   1275
      End
      Begin SICMACT.ActXCodCta ActXCodCta1 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   661
         Texto           =   "Cuenta N°:"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
   Begin VB.Frame Frame3 
      Height          =   675
      Left            =   60
      TabIndex        =   1
      Top             =   2430
      Width           =   6105
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "Aceptar"
         Height          =   345
         Left            =   180
         TabIndex        =   4
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   345
         Left            =   1410
         TabIndex        =   3
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   345
         Left            =   4830
         TabIndex        =   2
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1725
      Left            =   60
      TabIndex        =   0
      Top             =   720
      Width           =   6105
      Begin MSComctlLib.ListView LstAmpliado 
         Height          =   1560
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   2752
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Crédito"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Cliente"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Tipo Producto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Moneda"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Monto"
            Object.Width           =   1411
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmCredAmpliado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cCtaCod As String
Private MatCalend As Variant
Private rs_Ampliado As ADODB.Recordset
Public cPersCod As String
Public cMoneda As String
Public nIdCampana As Integer

Private Sub ActXCodCta1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdBuscar.SetFocus
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sNumTar As String
    Dim sClaveTar As String
    Dim nErr As Integer
    Dim sCaption As String
    Dim nEstado  As CaptacTarjetaEstado
    Dim nProducto As Integer
    
    If KeyCode = vbKeyF12 And ActXCodCta1.Enabled = True Then 'F12
        Dim bRetSinTarjeta As Boolean
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(nProducto, bRetSinTarjeta)
        If sCuenta <> "" Then
            ActXCodCta1.NroCuenta = sCuenta
            ActXCodCta1.SetFocusCuenta
        End If
    End If
    
End Sub


Sub InicializarRecord()
    Set rs_Ampliado = New ADODB.Recordset
    With rs_Ampliado.Fields
        .Append "cCtaCod", adVarChar, 18
        .Append "TipoProducto", adVarChar, 30
        .Append "dFecha", adDate
        .Append "nMonto", adDouble
        .Append "Moneda", adVarChar, 10
    End With
    rs_Ampliado.Open
    
End Sub

Private Sub ActXCodCta1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 123 Then 'pulso  la tecla f12
        FrmBuscadorCreditos.cPersCod = cPersCod
        FrmBuscadorCreditos.Show vbModal
        If cCtaCod <> "" And Len(cCtaCod) = 18 Then
            ActXCodCta1.CMAC = Mid(cCtaCod, 1, 3)
            ActXCodCta1.Age = Mid(cCtaCod, 4, 2)
            ActXCodCta1.Prod = Mid(cCtaCod, 6, 3)
            ActXCodCta1.Cuenta = Mid(cCtaCod, 9, 10)
        End If
    End If
End Sub

Private Sub CmdAceptar_Click()
    Dim i As Integer

    If sValidaDatos <> "" Then
        MsgBox sValidaDatos, vbInformation, "AVISO"
    Else

        If MsgBox("Desea establecer como un credito ampliado", vbInformation + vbYesNo) = vbYes Then
            Call InicializarRecord
           
           For i = 1 To LstAmpliado.ListItems.Count
                With rs_Ampliado
                    rs_Ampliado.AddNew
                    rs_Ampliado(0) = LstAmpliado.ListItems(i)
                    'ALPA 20100607******************************************************
                    'rs_Ampliado(1) = LstAmpliado.ListItems(i).SubItems(2)
                    rs_Ampliado(1) = Mid(LstAmpliado.ListItems(i).SubItems(2), 1, 30)
                    '*******************************************************************
                    rs_Ampliado(2) = gdFecSis
                    rs_Ampliado(3) = LstAmpliado.ListItems(i).SubItems(4)
                    rs_Ampliado(4) = LstAmpliado.ListItems(i).SubItems(3)
                    
                    .Update
                End With
            Next i
            Set frmCredSolicitud.rsAmpliado = Nothing
            Set frmCredSolicitud.rsAmpliado = rs_Ampliado
            Set rs_Ampliado = Nothing
            Unload Me
        End If
    End If
End Sub

Private Sub cmdBuscar_Click()
    'ARCV 09-03-2007
    If cMoneda <> Mid(Me.ActXCodCta1.NroCuenta, 9, 1) Then
        MsgBox "Debe seleccionar creditos de la misma moneda", vbInformation, "Mensaje"
        Exit Sub
    End If
    '----
    If Len(Me.ActXCodCta1.NroCuenta) = 18 And Me.ActXCodCta1.Texto <> "" Then
        Call CargarDatos
    End If
End Sub

Private Sub cmdCancelar_Click()

        InicializarRecord
        LstAmpliado.ListItems.Clear
        Me.ActXCodCta1.Cuenta = ""
        Me.ActXCodCta1.NroCuenta = ""
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Sub InicializarControles()
    CmdAceptar.Enabled = False
    CmdCancelar.Enabled = False
    ActXCodCta1.CMAC = gsCodCMAC
    ActXCodCta1.Age = gsCodAge
    LstAmpliado.ListItems.Clear
End Sub

Private Sub Form_Load()
    Call InicializarControles
End Sub

Sub CargarDatos()
    Dim rs As ADODB.Recordset
    Dim oNegCredito As COMNCredito.NCOMCredito
    Dim oAmpliado As COMDCredito.DCOMAmpliacion
    Dim nInteresFecha As Double
    Dim nMontoFecha As Double
    
    Dim Item As ListItem
    
    Set oNegCredito = New COMNCredito.NCOMCredito
    MatCalend = oNegCredito.RecuperaMatrizCalendarioPendiente(Me.ActXCodCta1.NroCuenta)
    If Not IsArray(MatCalend) Then
        MsgBox "La Cuenta no tiene Calendario pendiente", vbInformation, "Mensaje"
        Set oNegCredito = Nothing
        Exit Sub
    End If
    If UBound(MatCalend) = 0 Then
        MsgBox "La Cuenta no tiene Calendario pendiente", vbInformation, "Mensaje"
        Set oNegCredito = Nothing
        Exit Sub
    End If
    'MAVM 14092010 Se incluyo el Int Grac
    'nInteresFecha = oNegCredito.MatrizInteresGastosAFecha(Me.ActXCodCta1.NroCuenta, MatCalend, gdFecSis, True)
    nInteresFecha = oNegCredito.MatrizInteresGastosAFecha(Me.ActXCodCta1.NroCuenta, MatCalend, gdFecSis, True) + oNegCredito.MatrizInteresGraAFecha(Me.ActXCodCta1.NroCuenta, MatCalend, gdFecSis)
    nMontoFecha = oNegCredito.MatrizCapitalAFecha(Me.ActXCodCta1.NroCuenta, MatCalend)
    
    Set oNegCredito = Nothing
    
    nMontoFecha = nInteresFecha + nMontoFecha
    
    Set oAmpliado = New COMDCredito.DCOMAmpliacion
    Set rs = oAmpliado.ListaDatosAmpliacion(Me.ActXCodCta1.NroCuenta)
    Set oAmpliado = Nothing
    
    'LstAmpliado.ListItems.Clear
    If Not rs.BOF And Not rs.EOF Then
         Set Item = LstAmpliado.ListItems.Add(, , ActXCodCta1.NroCuenta)
         Item.SubItems(1) = IIf(IsNull(rs!Nombre), "", rs!Nombre)
         'ALPA 20100607 B2*******************************************************
         'Item.SubItems(2) = IIf(IsNull(rs!TipoProducto), "", rs!TipoProducto)
         Item.SubItems(2) = Mid(IIf(IsNull(rs!TipoProducto), "", rs!TipoProducto), 1, 30)
         '***********************************************************************
         Item.SubItems(3) = IIf(IsNull(rs!Moneda), "", rs!Moneda)
         Item.SubItems(4) = Format(nMontoFecha, "#0.00")
         
        nIdCampana = rs!IdCampana
        CmdAceptar.Enabled = True
        CmdCancelar.Enabled = True
    End If
    Set rs = Nothing
End Sub

Function sValidaDatos() As String
    Dim bValida As Boolean
    Dim oAmpliado As COMDCredito.DCOMAmpliacion
    

    'validamos que no exista ese registro en la tabla
    Set oAmpliado = New COMDCredito.DCOMAmpliacion
    'By Capi Acta 038-2007 punto 1
    'bValida = oAmpliado.VerificaAmpliado(Me.ActXCodCta1.NroCuenta)
    bValida = True
    Set oAmpliado = Nothing
    If bValida = False Then
        sValidaDatos = "Ya existe el credito como ampliado"
        Exit Function
    End If
    
    Set oAmpliado = New COMDCredito.DCOMAmpliacion
    bValida = oAmpliado.ValidacionCredito(Me.ActXCodCta1.NroCuenta)
    If bValida = False Then
        'sValidaDatos = "El credito tiene un estado no correcto"
        '***************RECO 20130913 INC: INC1309120009**************
        sValidaDatos = "No se puede ampliar el crédito por que tiene estado:" & UCase(oAmpliado.ObtenerEstadoProducto(Me.ActXCodCta1.NroCuenta))
        '***************END RECO**************************************

    End If
    
         
    ' Validamos que el credito este en esta vigente
    
End Function




