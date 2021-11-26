VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDISicmact 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema de Recuperaciones"
   ClientHeight    =   10365
   ClientLeft      =   1500
   ClientTop       =   -1125
   ClientWidth     =   16200
   Icon            =   "mdisicmact.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pbxFondo 
      Align           =   1  'Align Top
      Height          =   9375
      Left            =   0
      Picture         =   "mdisicmact.frx":030A
      ScaleHeight     =   9315
      ScaleWidth      =   28500
      TabIndex        =   2
      Top             =   600
      Width           =   28560
      Begin VB.Image Image1 
         Height          =   13500
         Left            =   0
         Picture         =   "mdisicmact.frx":51586
         Top             =   0
         Width           =   21600
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":7843F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":78759
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":78A73
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":78D8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":790A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":793C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":79553
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":7986D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":7A8BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":7B911
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":7C963
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":7D9B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":7EA07
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TlbMenu 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   28560
      _ExtentX        =   50377
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdImpre"
            Object.ToolTipText     =   "Impresora"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CmdCalc"
            Object.ToolTipText     =   "Calculadora"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CmdPagos"
            Object.ToolTipText     =   "Simulador de Pagos"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CmdPlazo"
            Object.ToolTipText     =   "Simulador de Plazo Fijo"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CmdCliente"
            Object.ToolTipText     =   "Posición Cliente"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdMantenimiento"
            Object.ToolTipText     =   "Mantenimiento Permisos"
            ImageIndex      =   13
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   120
      Top             =   720
   End
   Begin MSComctlLib.StatusBar SBBarra 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   14985
      Width           =   28560
      _ExtentX        =   50377
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   11994
            MinWidth        =   11994
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   "10/07/2019"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   "08:19"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   "NÚM"
         EndProperty
      EndProperty
   End
   Begin VB.Menu M0100000000 
      Caption         =   "&Archivo"
      Index           =   0
      Begin VB.Menu M0101000000 
         Caption         =   "Configurar &Impresora"
         Index           =   0
      End
      Begin VB.Menu M0101000000 
         Caption         =   "&Caracteres de Impresión"
         Index           =   1
      End
      Begin VB.Menu M0101000000 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu M0101000000 
         Caption         =   "&Salir"
         Index           =   3
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu M0300000000 
      Caption         =   "Colocacio&nes"
      Index           =   0
      Begin VB.Menu M0301000000 
         Caption         =   "&Recuperaciones"
         Index           =   4
         Begin VB.Menu M0301050000 
            Caption         =   "Proceso &Judicial"
            Index           =   1
            Begin VB.Menu M0301050100 
               Caption         =   "&Registro de Expedientes"
               Index           =   0
            End
            Begin VB.Menu M0301050100 
               Caption         =   "&Actuaciones Procesales"
               Index           =   1
            End
         End
         Begin VB.Menu M0301050000 
            Caption         =   "&Gastos de Recuperacion"
            Index           =   2
         End
         Begin VB.Menu M0301050000 
            Caption         =   "Actualizacion Metodo de &Liquidacion"
            Index           =   3
         End
         Begin VB.Menu M0301050000 
            Caption         =   "&Relaciones de Crédito"
            Index           =   4
            Begin VB.Menu M0301050200 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0301050200 
               Caption         =   "&Consulta"
               Index           =   1
            End
            Begin VB.Menu M0301050200 
               Caption         =   "Comision &Abogado"
               Index           =   2
            End
         End
         Begin VB.Menu M0301050000 
            Caption         =   "&Cancelación Créditos con Pago Judicial"
            Index           =   6
         End
         Begin VB.Menu M0301050000 
            Caption         =   "Cas&tigar Credito"
            Index           =   7
         End
         Begin VB.Menu M0301050000 
            Caption         =   "Consulta"
            Index           =   8
            Begin VB.Menu M0301050300 
               Caption         =   "Historial en Recuperaciones"
               Index           =   0
            End
            Begin VB.Menu M0301050300 
               Caption         =   "Consulta Pago a Gestores"
               HelpContextID   =   1
               Index           =   1
            End
         End
         Begin VB.Menu M0301050000 
            Caption         =   "Reportes"
            Index           =   9
            Begin VB.Menu M0301050400 
               Caption         =   "Reportes Mensuales"
               Index           =   0
            End
         End
         Begin VB.Menu M0301050000 
            Caption         =   "&Negociaciones"
            Index           =   12
            Begin VB.Menu M0301050500 
               Caption         =   "&Simulador"
               Index           =   0
            End
            Begin VB.Menu M0301050500 
               Caption         =   "&Registrar"
               Index           =   1
            End
            Begin VB.Menu M0301050500 
               Caption         =   "&Anular"
               Index           =   2
            End
         End
         Begin VB.Menu M0301050000 
            Caption         =   "&Vistos de Recuperaciones"
            Index           =   13
         End
         Begin VB.Menu M0301050000 
            Caption         =   "Bienes Adjudicados/Embargados/Vendidos"
            Index           =   16
         End
         Begin VB.Menu M0301050000 
            Caption         =   "Bienes Embargados"
            Index           =   17
         End
         Begin VB.Menu M0301050000 
            Caption         =   "Bloqueo Pago Recuperaciones"
            Index           =   18
         End
         Begin VB.Menu M0301050000 
            Caption         =   "Registro de Visita de Gestores"
            Enabled         =   0   'False
            Index           =   19
            Visible         =   0   'False
         End
         Begin VB.Menu M0301050000 
            Caption         =   "Campaña de Recuperaciones"
            Index           =   20
            Begin VB.Menu M0301052100 
               Caption         =   "Configuración"
               Index           =   0
            End
            Begin VB.Menu M0301052100 
               Caption         =   "Acoger Crédito"
               Index           =   1
            End
            Begin VB.Menu M0301052100 
               Caption         =   "Autorizar Operación"
               Index           =   2
            End
         End
         Begin VB.Menu M0301050000 
            Caption         =   "Creditos Transferidos"
            Index           =   21
            Begin VB.Menu M0301050600 
               Caption         =   "Distribución Pagos Focmac"
               Index           =   0
            End
         End
      End
   End
   Begin VB.Menu M0600000000 
      Caption         =   "&Sistema"
      Index           =   0
      Begin VB.Menu M0601000000 
         Caption         =   "Mantenimiento Permisos - &Responsabilidades"
         Index           =   1
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Mantenimiento &Permisos"
         Index           =   2
      End
      Begin VB.Menu M0701000000 
         Caption         =   "-"
         Index           =   2
      End
   End
   Begin VB.Menu M0800000000 
      Caption         =   "Herra&mientas"
      Index           =   0
      Begin VB.Menu M0801000000 
         Caption         =   "Editor de Textos"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu M0801000000 
         Caption         =   "Spooler de Impresión"
         Index           =   1
      End
      Begin VB.Menu M0801000000 
         Caption         =   "Configuración de Periféricos"
         Index           =   2
      End
      Begin VB.Menu M0801000000 
         Caption         =   "Cargar Logo &Penware"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu M0801000000 
         Caption         =   "Configuración de PINPADS"
         Index           =   4
      End
      Begin VB.Menu M0801000000 
         Caption         =   "Mensajes de Seguridad"
         Index           =   5
      End
   End
   Begin VB.Menu M0900000000 
      Caption         =   "A&yuda"
      Index           =   0
      Begin VB.Menu M0901000000 
         Caption         =   "&Contenido"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu M0901000000 
         Caption         =   "&Indice"
         Index           =   1
      End
      Begin VB.Menu M0901000000 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu M0901000000 
         Caption         =   "&Acerca del Sistema..."
         Index           =   3
      End
   End
End
Attribute VB_Name = "MDISicmact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Api para Ejecutar Internet Explorer
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpoperation As String, ByVal lpfile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowcmd As Long) As Long


Private Sub cmdVer_Click()
      'FrmCapAutOpeEstados.Show
End Sub

Private Sub Command1_Click()

Dim loPrevio As New previo.clsprevio
Dim lscadimp As String
Dim i As Integer

i = 0

Do While MsgBox("Imprimir???", vbYesNo, "Aviso") = vbYes
    
    lscadimp = ""
    lscadimp = Chr$(27) & Chr$(64)
    lscadimp = lscadimp & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
    lscadimp = lscadimp & Chr$(27) & Chr$(15) 'Condensada
    lscadimp = lscadimp & Chr$(27) & Chr$(67) & Chr$(22) 'Longitud de página a 22 líneas'
    lscadimp = lscadimp & Chr$(27) & Chr$(77)  'Tamaño 10 cpi
    lscadimp = lscadimp & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
    
    lscadimp = lscadimp & Chr$(27) & Chr$(103)
    lscadimp = lscadimp & "   " & Chr(10)
    lscadimp = lscadimp & "   " & Chr(10)
    lscadimp = lscadimp & Chr$(27) & Chr$(77)
    
    lscadimp = lscadimp & Chr$(27) & Chr$(69) 'activa negrita
    lscadimp = lscadimp & Chr$(27) + Chr$(72) ' desactiva negrita
     
    If i > 0 Then
        lscadimp = lscadimp & "" & Chr(10)
        lscadimp = lscadimp & "" & Chr(10)
        lscadimp = lscadimp & "" & Chr(10)
        lscadimp = lscadimp & "" & Chr(10)
    End If
     
     
    lscadimp = lscadimp & "" & Chr(10)
    lscadimp = lscadimp & "" & Chr(10)
    lscadimp = lscadimp & "" & Chr(10)
    lscadimp = lscadimp & "" & Chr(10)
    lscadimp = lscadimp & "" & Chr(10)
    lscadimp = lscadimp & "" & Chr(10)
    lscadimp = lscadimp & Space(28) & "89123451231231321313212316789" & Space(26) & "1234567890" & Space(38) & "3456789123123123132" & Chr(10)
    lscadimp = lscadimp & Space(83) & "1234567890" & Chr(10)
    lscadimp = lscadimp & "" & Chr(10)
    lscadimp = lscadimp & Space(38) & "Apellidos" & Chr(10)
    lscadimp = lscadimp & Space(38) & "Nombre" & Chr(10)
    lscadimp = lscadimp & Space(38) & "DNI123456789123456789123456789123456789123456789123456789213456789123456789" & Space(24) & "9999" & Chr(10)
    lscadimp = lscadimp & Space(38) & "Calle" & Chr(10)
    lscadimp = lscadimp & Space(38) & "Urbanizacion" & Chr(10)
    lscadimp = lscadimp & Space(38) & "Ciudad" & Chr(10)
    lscadimp = lscadimp & Space(38) & "Telefono" & Chr(10)
    lscadimp = lscadimp & "" & Chr(10)
    lscadimp = lscadimp & Space(38 + 79) & "3456789123456789123 999" & Chr(10)
    lscadimp = lscadimp & Space(38 + 79) & "3456789123456789123 999" & Chr(10)
    lscadimp = lscadimp & Space(38 + 79) & "3456789123456789123 999" & Chr(10)
    lscadimp = lscadimp & Space(38 + 79) & "3456789123456789123 999" & Chr(10)
    lscadimp = lscadimp & "" & Chr(10)
    lscadimp = lscadimp & Space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Space(22) & "3456789123456789123 999" & Chr(10)
    lscadimp = lscadimp & Space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Space(22) & "3456789123456789123 999" & Chr(10)
    lscadimp = lscadimp & Space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lscadimp = lscadimp & Space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Space(22) & "3456789123456789123 999" & Chr(10)
    lscadimp = lscadimp & Space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Space(22) & "3456789123456789123 999" & Chr(10)
    lscadimp = lscadimp & Space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Space(22) & "3456789123456789123 999" & Chr(10)
    lscadimp = lscadimp & Space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lscadimp = lscadimp & Space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lscadimp = lscadimp & Space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lscadimp = lscadimp & Space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Space(22) & "3456789123456789123 999" & Chr(10)
    lscadimp = lscadimp & Space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lscadimp = lscadimp & Space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lscadimp = lscadimp & Space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lscadimp = lscadimp & Space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lscadimp = lscadimp & Space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lscadimp = lscadimp & Space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lscadimp = lscadimp & Space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lscadimp = lscadimp & Space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lscadimp = lscadimp & Space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lscadimp = lscadimp & Space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lscadimp = lscadimp & "" & Chr(10)
    lscadimp = lscadimp & Space(30) & "Inc. 11: pasados 30 dias del vencimiento de su contrato" & Chr(10)
    lscadimp = lscadimp & Space(30) & "sus joyas entrarán a Remate Sin Notificar" & Chr(10)
    lscadimp = lscadimp & "" & Chr(10)
    lscadimp = lscadimp & "" & Chr(10)
    lscadimp = lscadimp & "" & Chr(10)
    lscadimp = lscadimp & "" & Chr(10)
    lscadimp = lscadimp & "" & Chr(10)
    lscadimp = lscadimp & "" & Chr(10)
    lscadimp = lscadimp & "" & Chr(10)
    lscadimp = lscadimp & "" & Chr(10)
    lscadimp = lscadimp & "" & Chr(10)
    lscadimp = lscadimp & "" & Chr(10)
    lscadimp = lscadimp & "" & Chr(10)
    lscadimp = lscadimp & "" & Chr(10)
    lscadimp = lscadimp & Space(116) & "123456789123456789123456789" & Chr(10)
    lscadimp = lscadimp & Space(20) & "12345678912345678912345678912345678912345678912345678912345678912345678912" & Space(2) & "5678912345" & Space(19) & "123456789" & Chr(10)
    lscadimp = lscadimp & Chr$(27) + Chr$(18) ' cancela condensada

    loPrevio.PrintSpool sLpt, lscadimp, False
    
    i = i + 1
    
Loop

        
End Sub

Private Sub M0101000000_Click(Index As Integer)
    Select Case Index
        Case 0
            frmImpresora.Show 1
        Case 1
            frmCaracImpresion.Show 1
        Case 3
            If MsgBox("Esta Seguro que Desea Salir del Sistema ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                Call SalirSICMACMNegocio
                End
            End If
    End Select
End Sub

'yihu
Private Sub M0301050000_Click(Index As Integer)
    Select Case Index
        'Case 0 ' Ingreso a Recup de Otras Entidades
        '    frmColRecIngresoOtrasEnt.Show 1
        Case 2 ' Gastos en Recuperaciones
            frmColRecGastosRecuperaciones.Show 1
        Case 3 ' Metodo de Liquidacion
            frmColRecMetodoLiquid.Show 1
        Case 5 ' Pago
            frmColRecPagoCredRecup.Inicio gColRecOpePagJudSDEfe, "PAGO CREDITO EN RECUPERACIONES", gsCodCMAC, gsNomCmac, True
         '** Juez 20120418 **************************
        Case 6 ' Cancelacion
            'frmColRecCancelacion.Show 1
            frmColRecCancPagoJudicial.Show 1
        '** End Juez *******************************
        Case 7 ' Castigo
            frmColRecCastigar.Show 1
        Case 8
        
        Case 10
            'frmGarLevant.Show 1'EJVG20160310-> Se comentó este Levantamiento xq por la misma opción de Garantías se realizará.
        Case 11
            'frmGarantExtorno.Show 1'EJVG20160310-> Se comentó este Levantamiento xq por la misma opción de Garantías se realizará.
        Case 13
            FrmColRecVistoRecup.Show 1
        Case 14
            'frmCredTransfRecupeGarant.Show 1
        Case 15
            'frmCredTransfGarantiaAdjudiSaneado.Show 1
        Case 16
            frmColBienesAdjudLista.Show 1
        Case 17
            'frmColEmbargadosListar.Show 1
        'MADM 20111010
        Case 18
            FrmBloqueaRecupera.Show 1 'X Mem
        Case 19 '*** PEAC 20120816
            FrmColRecRegVisitaCliente.Show 1
            
    End Select
End Sub
'yihu
Private Sub M0301050100_Click(Index As Integer)
    Select Case Index
        Case 0
            frmColRecExped.Show 1
        Case 1
            frmColRecActuacionesProc.Inicia "N"
    End Select
End Sub

Private Sub M0301050200_Click(Index As Integer)
Dim oCredRel As New UCredRelac_Cli 'COMDCredito.UCOMCredRela   'UCredRelacion
Select Case Index
    Case 0
        frmCredRelaCta.Inicio oCredRel, InicioMantenimiento, InicioRegistroForm
        Set oCredRel = Nothing
    Case 1
        frmCredRelaCta.Inicio oCredRel, InicioMantenimiento, InicioConsultaForm
        Set oCredRel = Nothing
    Case 2
        frmColRecComision.Show 1
End Select
End Sub

'yihu
Private Sub M0301050300_Click(Index As Integer)
Select Case Index
    Case 0
        frmColRecRConsulta.Inicia "Consulta de Pagos de Créditos Judiciales"
    Case 1
        FrmColRecPagGestor.Show vbModal
End Select
End Sub
'yihu
Private Sub M0301050400_Click(Index As Integer)
Select Case Index
    Case 0
        frmColRecReporte.Inicia "Reportes de Recuperaciones"
End Select
End Sub
'yihu
Private Sub M0301050500_Click(Index As Integer)
Select Case Index
    Case 0 ' Simulador
        Call frmColRecNegCalculaCalendario.Inicio(1)
    Case 1 'Registrar Negociacion
        frmColRecNegRegistro.Inicia (True)
    Case 2 'Resolver Negociacion
        frmColRecNegRegistro.Inicia (False)
End Select

End Sub
'yihu
'FRHU 20150428 ERS022-2015
Private Sub M0301050600_Click(Index As Integer)
    Select Case Index
        Case 0
            frmColTransfCancelPago.Show 1
    End Select
End Sub
'FIN FRHU 20150428

'yihu
'WIOR 20150602 ***
Private Sub M0301052100_Click(Index As Integer)
Select Case Index
    Case 0
        frmRecupCampConfig.Show 1
    Case 1
        frmRecupCampAcoger.Show 1
    Case 2
        frmRecupCampAuto.Show 1
End Select
End Sub
'WIOR FIN ********

Public Function RepHavDevBoveda(psFecSis As Date, psNomCmac As String, psNomAge As String, psCodAge As String) As String
  Dim cMovNro As String, rstemp As ADODB.Recordset
  Dim oRep As nCaptaReportes
  Dim xlAplicacion As Excel.Application
  Dim xlLibro As Excel.Workbook
  Dim xlHoja1 As Excel.Worksheet
  Dim nFila As Long, i As Long
  Dim MONHAB  As Double, MONDEV As Double
  Dim NUMHAB  As Double, NUMDEV As Double
  Dim lsArchivoN As String, lbLibroOpen As Boolean

  MONHAB = 0
  MONDEV = 0
  NUMHAB = 0
  NUMDEV = 0
    
  Set oRep = New nCaptaReportes
    Set rstemp = oRep.RepHavDevBoveda(Format(psFecSis, "yyyymmdd"), psCodAge)
   
  If rstemp.EOF Then
    MsgBox "No se encontro información para este reporte", vbOKOnly + vbInformation, "Aviso"
    Exit Function
  End If
   
  lsArchivoN = App.Path & "\Spooler\RepHABDEVBOV" & Format(gdFecSis & " " & Time, "yyyymmddhhmmss") & gsCodUser & ".xls"
   
  'OleExcel.Class = "ExcelWorkSheet"
  lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
  If lbLibroOpen Then
            Set xlHoja1 = xlLibro.Worksheets(1)
            ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
                                   
            nFila = 1
            
            xlHoja1.Cells(nFila, 1) = gsNomCmac
            nFila = 2
            xlHoja1.Cells(nFila, 1) = gsNomAge
            xlHoja1.Range("F2:H2").MergeCells = True
            xlHoja1.Cells(nFila, 6) = Format(gdFecSis, "Long Date")
             
'             prgBar.value = 2
                              
            nFila = 3
            xlHoja1.Cells(nFila, 1) = "REPORTE DE HABILITACIONES Y DEVOLUCIONES PARA BOVEDA " & psNomAge
                         
            xlHoja1.Range("A1:M3").Font.Bold = True
            
            xlHoja1.Range("A3:M3").MergeCells = True
            xlHoja1.Range("A3:A3").HorizontalAlignment = xlCenter
             
            'xlHoja1.Range("A5:H5").AutoFilter
            
            nFila = 5
            
                nFila = nFila + 1
                
            xlHoja1.Range("A" & nFila & ":E" & nFila).Font.Bold = True
            xlHoja1.Range("A" & nFila & ":E" & nFila).MergeCells = True
            xlHoja1.Range("A" & nFila & ":E" & nFila).HorizontalAlignment = xlCenter
            xlHoja1.Cells(nFila, 1) = "HABILITACIONES"
            
              nFila = nFila + 1
            xlHoja1.Cells(nFila, 1) = "ITEM"
            xlHoja1.Cells(nFila, 2) = "MONEDA"
            xlHoja1.Cells(nFila, 3) = "IMPORTE"
            xlHoja1.Cells(nFila, 4) = "USUARIO"
            xlHoja1.Cells(nFila, 5) = "NOMBRE USUARIO"
            xlHoja1.Cells(nFila, 6) = "FECHA"
            xlHoja1.Cells(nFila, 7) = "HORA"
            
            i = 0
            While Not rstemp.EOF
            
               If rstemp.Fields("COPECOD") <> 901017 Then
                  GoTo Men
               End If
               
                nFila = nFila + 1
                
'                prgBar.value = ((i) / RSTEMP.RecordCount) * 100
                
                i = i + 1
                
                xlHoja1.Cells(nFila, 1) = Format(i, "0000")
                xlHoja1.Cells(nFila, 2) = rstemp!nmoneda
                xlHoja1.Cells(nFila, 3) = Format(rstemp!nMovImporte, "#0.00")
                xlHoja1.Cells(nFila, 4) = rstemp!CUSUDEST
                xlHoja1.Cells(nFila, 5) = rstemp!Nombre
                xlHoja1.Cells(nFila, 6) = Format(CDate(Mid(rstemp!cMovNro, 5, 2) & "-" & Mid(rstemp!cMovNro, 7, 2) & "-" & Left(rstemp!cMovNro, 4)), "dd/MM/yyyy")
                xlHoja1.Cells(nFila, 7) = Mid(rstemp!cMovNro, 9, 2) & ":" & Mid(rstemp!cMovNro, 11, 2) & ":" & Mid(rstemp!cMovNro, 13, 2)
                                            
                MONHAB = MONHAB + rstemp!nMovImporte
                
                rstemp.MoveNext
                
            Wend
                         
Men:

            NUMHAB = i
            
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 1) = "TOTAL: " & CStr(NUMHAB)
            xlHoja1.Cells(nFila, 3) = Format(MONHAB, "#0.00")

            nFila = nFila + 2
                
            xlHoja1.Range("A" & nFila & ":E" & nFila).Font.Bold = True
            xlHoja1.Range("A" & nFila & ":E" & nFila).MergeCells = True
            xlHoja1.Range("A" & nFila & ":E" & nFila).HorizontalAlignment = xlCenter
            
            xlHoja1.Cells(nFila, 1) = "DEVOLUCIONES"
                
                nFila = nFila + 1
            xlHoja1.Cells(nFila, 1) = "ITEM"
            xlHoja1.Cells(nFila, 2) = "MONEDA"
            xlHoja1.Cells(nFila, 3) = "IMPORTE"
            xlHoja1.Cells(nFila, 4) = "USUARIO"
            xlHoja1.Cells(nFila, 5) = "NOMBRE USUARIO"
            xlHoja1.Cells(nFila, 6) = "FECHA"
            xlHoja1.Cells(nFila, 7) = "HORA"
            
            i = 0
            While Not rstemp.EOF

                nFila = nFila + 1
                
'                prgBar.value = ((i) / RSTEMP.RecordCount) * 100
                
                i = i + 1
                
                xlHoja1.Cells(nFila, 1) = Format(i, "0000")
                xlHoja1.Cells(nFila, 2) = rstemp!nmoneda
                xlHoja1.Cells(nFila, 3) = Format(rstemp!nMovImporte, "#0.00")
                xlHoja1.Cells(nFila, 4) = rstemp!CUSUDEST
                xlHoja1.Cells(nFila, 5) = rstemp!Nombre
                xlHoja1.Cells(nFila, 6) = Format(CDate(Mid(rstemp!cMovNro, 5, 2) & "-" & Mid(rstemp!cMovNro, 7, 2) & "-" & Left(rstemp!cMovNro, 4)), "dd/MM/yyyy")
                xlHoja1.Cells(nFila, 7) = Mid(rstemp!cMovNro, 9, 2) & ":" & Mid(rstemp!cMovNro, 11, 2) & ":" & Mid(rstemp!cMovNro, 13, 2)
                                            
                MONDEV = MONDEV + rstemp!nMovImporte
                
                rstemp.MoveNext
                
            Wend
            
            NUMDEV = i
            
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 1) = "TOTAL: " & CStr(NUMDEV)
            xlHoja1.Cells(nFila, 3) = Format(MONDEV, "#0.00")
            
             Set rstemp = New Recordset
            
            Set rstemp = oRep.REPBOVSALDOS(psCodAge, Format(psFecSis, "yyyymmdd"), Format(psFecSis, "yyyymmdd"))
            
               nFila = nFila + 1
               
             xlHoja1.Range("A" & nFila & ":E" & nFila).Font.Bold = True
             xlHoja1.Range("A" & nFila & ":E" & nFila).MergeCells = True
             xlHoja1.Range("A" & nFila & ":E" & nFila).HorizontalAlignment = xlCenter
           
                xlHoja1.Cells(nFila, 1) = "SALDOS FINALES"
                
                nFila = nFila + 1
                xlHoja1.Cells(nFila, 1) = "USUARIO"
                xlHoja1.Cells(nFila, 2) = "NOMBRE USUARIO"
                xlHoja1.Cells(nFila, 3) = "MONTO S/."
                xlHoja1.Cells(nFila, 4) = "MONTO U$."
                xlHoja1.Cells(nFila, 5) = "FECHA"
            
            While Not rstemp.EOF
                            
                nFila = nFila + 1

                xlHoja1.Cells(nFila, 1) = rstemp!cUser
                xlHoja1.Cells(nFila, 2) = rstemp!cPersNombre
                xlHoja1.Cells(nFila, 3) = rstemp!solesmonto
                xlHoja1.Cells(nFila, 4) = rstemp!dolaresmonto
                xlHoja1.Cells(nFila, 5) = Format(CDate(Mid(rstemp!dfecha, 5, 2) & "/" & Right(rstemp!dfecha, 2) & "/" & Left(rstemp!dfecha, 4)), "dd/MM/yyyy")
                               
                    rstemp.MoveNext
            Wend
            
            xlHoja1.Columns.AutoFit
            
            xlHoja1.Cells.Select
            xlHoja1.Cells.Font.Name = "Arial"
            xlHoja1.Cells.Font.Size = 9
            xlHoja1.Cells.EntireColumn.AutoFit
                
            'Cierro...
            'OleExcel.Class = "ExcelWorkSheet"
            ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
            'OleExcel.SourceDoc = lsArchivoN
            'OleExcel.Verb = 1
            'OleExcel.Action = 1
            'OleExcel.DoVerb -1
            
'            prgBar.value = 100
            
   End If
   
   RepHavDevBoveda = "GENERADO"
   
   Set rstemp = Nothing
   
'   prgBar.Visible = False
  
End Function

 '***********************************************************
' Inicia Trabajo con EXCEL, crea variable Aplicacion y Libro
'***********************************************************
Private Function ExcelBegin(psArchivo As String, _
        xlAplicacion As Excel.Application, _
        xlLibro As Excel.Workbook, Optional pbBorraExiste As Boolean = True) As Boolean
        
Dim fs As New Scripting.FileSystemObject
On Error GoTo ErrBegin
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application

If fs.FileExists(psArchivo) Then
   If pbBorraExiste Then
      fs.DeleteFile psArchivo, True
      Set xlLibro = xlAplicacion.Workbooks.Add
   Else
      Set xlLibro = xlAplicacion.Workbooks.Open(psArchivo)
   End If
Else
   Set xlLibro = xlAplicacion.Workbooks.Add
End If
ExcelBegin = True
Exit Function
ErrBegin:
  MsgBox Err.Description, vbInformation, "Aviso"
  ExcelBegin = False
End Function

'***********************************************************
' Final de Trabajo con EXCEL, graba Libro
'***********************************************************
Private Sub ExcelEnd(psArchivo As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional plSave As Boolean = True)
On Error GoTo ErrEnd
   If plSave Then
        xlHoja1.SaveAs psArchivo
   End If
   xlLibro.Close
   xlAplicacion.Quit
   Set xlAplicacion = Nothing
   Set xlLibro = Nothing
   Set xlHoja1 = Nothing
Exit Sub
ErrEnd:
   MsgBox Err.Description, vbInformation, "Aviso"
End Sub

'********************************
' Adiciona Hoja a LibroExcel
'********************************
Private Sub ExcelAddHoja(psHojName As String, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet)
For Each xlHoja1 In xlLibro.Worksheets
    If xlHoja1.Name = psHojName Then
       xlHoja1.Delete
       Exit For
    End If
Next
Set xlHoja1 = xlLibro.Worksheets.Add
xlHoja1.Name = psHojName
End Sub

'yihu
Private Sub M0601000000_Click(Index As Integer)
    Select Case Index
        Case 0 'Parametros
            
        Case 1 'Permisos responsabilidades
            '**DAOR 20071122 ******************
             frmMantPermisosResponsable.Show 1
            '**********************************
        Case 2 'permisos
            frmMantPermisos.Show 1
    End Select
End Sub

'yihu
Private Sub M0801000000_Click(Index As Integer)
    Select Case Index
        'Herramientas
        Case 0
           ' frmCartaPrend.Show 1
        Case 1
            'frmSpooler.Show 1
        Case 2
            frmSetupCOM.Show 1
          '  frmComparaTablas.Show 1
        Case 3
            'frmExplorerSicmact.Show 1
            'frmControlCalidadSicmac.Show 1
        Case 4
            frmPITPinPadSeleccion.Show vbModal
        'WIOR 2013906 ********************
        Case 5
            'frmSegMensajes.Show 1
        'WIOR FIN ************************
    End Select
End Sub

'RECO FIN****************************************
Private Sub MDIForm_Click()
'  Form1.Show
End Sub

'yihu
'Private Sub MDIForm_Load()
' Me.Icon = LoadPicture(App.Path & gsRutaIcono)
'Timer1.Enabled = False
'CargaMensajes  'WIOR 20130826
'If gsCodCargo = "005002" Or gsCodCargo = "005003" Or gsCodCargo = "005004" Or gsCodCargo = "005005" Then
'
'    'VAPI INTEGRAR BLOQUEO SICMACM PARA ANALISTAS QUE USAN EL APLICATIVO MOVIL 20160523
'    Dim oDHojaRuta As DCOMhojaRuta
'    Set oDHojaRuta = New DCOMhojaRuta
'
'    If oDHojaRuta.participaHojaRuta(gsCodUser) Then
'
'        Dim nRespuesta As Integer
'        Dim bEsHoraLimite As Boolean
'        bEsHoraLimite = oDHojaRuta.esHoraLimite()
'        nRespuesta = oDHojaRuta.puedeGenerar(gsCodUser, bEsHoraLimite)
'        Dim cMensaje As String
'        Dim bEntrar As Boolean
'
'        Select Case nRespuesta
'            Case 1:
'                oDHojaRuta.SolicitarVistoHojaRuta gsCodUser, gsCodAge, nRespuesta
'                cMensaje = "Tiene resultados de días anteriores pendientes por enviar"
'            Case 2:
'                oDHojaRuta.SolicitarVistoHojaRuta gsCodUser, gsCodAge, nRespuesta
'                cMensaje = "No ha generado su hoja de ruta en días anteriores"
'            Case 3, 4:
'                bEntrar = True
'        End Select
'
'        If (oDHojaRuta.tieneVistoPendiente(gsCodUser)) Then
'            cMensaje = cMensaje & ",Solicita un visto bueno, regularice sus pendientes para poder entrar al sistema."
'            bEntrar = False
'        Else
'            bEntrar = True
'        End If
'
'        If Not bEntrar Then
'            MsgBox cMensaje, vbInformation, "AVISO"
'           End
'        End If
'    'FIN VAPI
'    End If
'
'End If

'FRHU20140319 RQ13874
'Dim objCred As New COMDCredito.DCOMCredito
'Dim objRS As ADODB.Recordset
'Dim valor As Integer
'Screen.MousePointer = 0
'Set objRS = objCred.ValidarCargoProyeccionColocAge(gsCodCargo)
'If objRS Is Nothing Then
'    valor = 0
'Else
'    If Not objRS.EOF And Not objRS.BOF Then
'        valor = objRS!valor
'    End If
'End If
'If valor = 1 Then
'    Call frmProyeccionPorAgencias.Inicio(gsCodAge, gdFecSis)
'End If
'Screen.MousePointer = 11
''FIN FRHU20140319 RQ13874
'
'End Sub

Private Sub AvisoOperacionesPendientes()

     
    Dim lsCadAux As String
    Dim lsSQL As String
    Dim coConex As DConecta
    Dim lsCola As String
    Dim lrst As ADODB.Recordset
    
    Set lrst = New ADODB.Recordset
    Set coConex = New DConecta
    coConex.AbreConexion

    
    Dim lrOperaciones As ADODB.Recordset
                        
                                                
     lsCola = " where (nAutEstado = " & gAhoEstAprobOpeAutorizado & _
        "  and (select count(nMovNro) from MovRef where MovRef.nMovNro = Mcao.nMovNro)=1 and  Mcao.cUserOri ='" & gsCodUser & "')" & _
        " or (nAutEstado = " & gAhoEstAprobOpeRechazado & " and    Mcao.cUserOri ='" & gsCodUser & "')  and cHabilitado ='S'"
        
'     lsCola = " where (nAutEstado = " & gAhoEstAprobOpeRechazado & " and    Mcao.cUserOri ='" & gsCodUser & "')"
'
'
'     lsCola = " where (nAutEstado = " & gAhoEstAprobOpeAutorizado & _
'        "  and (select count(nMovNro) from MovRef where MovRef.nMovNro = Mcao.nMovNro)=1 and  Mcao.cUserOri ='" & gsCodUser & "')"

                                    
    lsSQL = "Select cCtaCod,(select cPersNombre from Persona where Persona.cPerscod = Mcao.cPersCodCli)   as  cPersNombre," & _
    "(select cOpeDesc + replicate(' ',100) + OpeTpo.cOpeCod from OpeTpo where OpeTpo.cOpeCod = Mcao.cOpeCod  ) cOpeDes," & _
    "(select cOpeDesc from OpeTpo where OpeTpo.cOpeCod = Mcao.cOpeCodOri  ) cOpeDesOri," & _
    "(select case Mcao.nCodMon when 1 then 'S/.' + replicate(' ',100) + '1'   when 2 then '$.'+ replicate(' ',100) + '2' end) Monedas,Mcao.nMontoAprobado,  " & _
    "(select cConsDescripcion from constante where nConsCod = 9051  and  constante.nConsvalor = Mcao.nAutEstado) sAutEstadoDes," & _
    "Mcao.cAutObs, Mcao.cUltimaActualizacion UltAct, Mcao.nMovNro nMov , Mcao.cOpeCodOri OpeOri, Mcao.cUserOri UserOri " & _
    "From MovCapAutorizacionOpe  As Mcao" & lsCola
            
        
    Set lrst = coConex.CargaRecordSet(lsSQL)
    
    If lrst.EOF Then
        MsgBox "No Existen Operaciones Pendientes Autorizadas", vbInformation, "Aviso"
        Exit Sub
    End If
    
    'MsgBox lrst.RecordCount
    lsCadAux = ""
    
    lrst.MoveFirst
    Do While Not lrst.EOF
        'MsgBox lrst("cOpeDesOri")
        lsCadAux = lsCadAux + PstaNombre(Trim(lrst("cPersNombre")), True) + Space(10) + Trim(lrst("cOpeDesOri")) + Space(10) + Trim(lrst("sAutEstadoDes")) + Chr(13)
        lrst.MoveNext
    Loop
    MsgBox lsCadAux, vbInformation, "Operaciones Pendientes"


End Sub

Private Sub mnuDetalleOperacionesII_Click()
     Dim oReportCreditos As DReportCreditos
      Set oReportCreditos = New DReportCreditos
      Call oReportCreditos.ReporteDetalleOP(gsCodUser, gdFecSis)
      Set oReportCreditos = Nothing
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    '**DAOR 20090203, Regitro de salida del sistema****
    Call SalirSICMACMNegocio
    '**************************************************
End Sub

'Private Sub mnuRelGarant_Click()
'    FrmCredRelGarantias.Show vbModal
'End Sub

Private Sub Timer1_Timer()

End Sub
Private Sub TlbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "cmdImpre"
           frmImpresora.Show 1
        Case "CmdCalc"
           SCalculadora
        Case "CmdPagos"
            'frmCredCalendPagos.Show 1'WIOR 20150209 COMENTO
            'frmCredCalendPagos.InicioSim 'WIOR 20150209 'ARLO20180625 COMENTO
            frmCredCalendPagosNEW.InicioSim 'ARLO20180723 ERS037 -2018
        Case "CmdPlazo"
            frmCapSimulacionPF.Show 1
        Case "CmdCliente"
            frmPosicionCli.Show 1
        '->***** LUCV20190323, Según RO-1000373
        Case "cmdMantenimiento"
            frmMantPermisos.Show 1
        '<-Fin LUCV20190323*****
    End Select
End Sub

Public Sub SCalculadora()
    Dim valor, i
    valor = Shell("calc.exe", 1)  ' Ejecuta la Calculadora.
    'AppActivate valor             ' Activa la Calculadora. '->***** LUCV20190323, Comentó según RO-1000373

End Sub


'**DAOR 20090203 , Registro de salida del sistema SICMACM Negocio************
'Bitacora Version 201011
Sub SalirSICMACMNegocio()
Dim oSeguridad As New COMManejador.Pista
    'Call oSeguridad.InsertarPista(gIngresarSalirSistema, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gSalirSistema, "Salida del SICMACM Negocio" & " Versión: " & Format(App.Major, "#0") & "." & Format(App.Minor, "#0") & "." & Format(App.Revision, "#0") & "-" & frmLogin.gsFechaVersion) 'marg20181217 comento
    Call oSeguridad.InsertarPista(gIngresarSalirSistema, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gSalirSistema, "Salida del SICMACM Recuperaciones" & " Versión: " & Format(App.Major, "#0") & "." & Format(App.Minor, "#0") & "." & Format(App.Revision, "#0") & "-" & frmLogin.gsFechaVersion) 'marg20181217 agrego
     If oSeguridad.ValidaAccesoPistaRF(gsCodUser) Then
            'Call oSeguridad.InsertarPistaSesion(gIngresarSalirSistema, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gSalirSistema, 1) 'marg20181217 comento
            Call oSeguridad.InsertarPistaSesion(gIngresarSalirSistema, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gSalirSistema, 6) 'marg20181217 agrego
            'Call oSeguridad.ActualizarPistaSesion(gsCodPersUser, GetMaquinaUsuario, 1) 'JUEZ 20160125 'marg20181217 comento
            Call oSeguridad.ActualizarPistaSesion(gsCodPersUser, GetMaquinaUsuario, 6) 'JUEZ 20160125 'marg20181217 agrego
     End If
    Set oSeguridad = Nothing
End Sub
'****************************************************************************

'WIOR 20130826 *************************************************************
Private Sub CargaMensajes()
Dim oSeg As COMDPersona.UCOMAcceso
Dim rsSeg As ADODB.Recordset

Set oSeg = New COMDPersona.UCOMAcceso
Set rsSeg = oSeg.ObtenerMensajeSeguridad(, True)

If Not (rsSeg.EOF And rsSeg.BOF) Then
    'frmSegMensajeMostrar.Inicio (Trim(rsSeg!cMensaje))
End If

Set oSeg = Nothing
Set rsSeg = Nothing
End Sub
'WIOR FIN *******************************************************************

' *** RIRO SEGUN TI-ERS108-2013 ***
Public Function VerificarRFIII() As Boolean
    Dim rsRF3 As New ADODB.Recordset
    Dim sMensaje As String
    Set rsRF3 = ValidarRFIII
    sMensaje = ""
    If Not (rsRF3.EOF Or rsRF3.BOF) And rsRF3.RecordCount > 0 Then
        If gsCodCargo = "006005" Then ' *** SI ES "SUPERVISOR"
            If Not rsRF3!bOpcionesSimultaneas And rsRF3!bModoSupervisor Then
                sMensaje = "Actualmente no puede realizar ninguna operacioón debido a que el RFIII se encuentra activo en modo supervisor " & vbNewLine & _
                "por favor active el segundo perfil del RFIII para poder acceder a esta opción "
            End If
        ElseIf gsCodCargo = "007026" Then ' *** SI ES "RFIII"
            If rsRF3!bPerfilCambiado Then
                If rsRF3!bModoSupervisor Then
                    sMensaje = "Su perfil ha cambiado a modo supervisor, debe volver a acceder al sistema"
                Else
                    sMensaje = "Su perfil ha cambiado a modo normal, debe volver a acceder al sistema"
                End If
            End If
        End If
    End If
    If Len(sMensaje) > 0 Then
        MsgBox sMensaje, vbExclamation, "Aviso"
        If gsCodCargo = "007026" Then ' *** SI ES "RFIII"
           VerificarRFIII = False
           End
        ElseIf gsCodCargo = "006005" Then
           VerificarRFIII = False
           Exit Function
        End If
    End If
    VerificarRFIII = True
End Function
' *** FIN RIRO ***
'RECO20151111 ERS061-2015 *************************
Private Function VerificaGrupoPermisoPostCierre() As Boolean
    Dim oCons As New COMDConstSistema.DCOMGeneral
    Dim sGrupoAutorizado As String
    Dim nGrupoTmp1 As String
    Dim nGrupoTmp2 As String
    Dim i As Integer
    Dim j As Integer
            
    sGrupoAutorizado = oCons.LeeConstSistema(516)
    VerificaGrupoPermisoPostCierre = False
    For i = 1 To Len(sGrupoAutorizado)
        If Not Mid(sGrupoAutorizado, i, 1) = "," Then
            nGrupoTmp1 = nGrupoTmp1 & Mid(sGrupoAutorizado, i, 1)
        Else
            For j = 1 To Len(gsGruposUser)
                If Not Mid(gsGruposUser, j, 1) = "," Then
                    nGrupoTmp2 = nGrupoTmp2 & Mid(gsGruposUser, j, 1)
                Else
                    If nGrupoTmp1 = nGrupoTmp2 Then
                        VerificaGrupoPermisoPostCierre = True
                        Exit Function
                    End If
                    nGrupoTmp2 = ""
                End If
            Next
            nGrupoTmp1 = ""
        End If
    Next
End Function
'RECO FIN *****************************************

'->***** LUCV20190323, Según RO-1000373
Private Sub MDIForm_Load()
    'Quita el borde de los dos controles
    Image1.BorderStyle = 0
    pbxFondo.BorderStyle = 0

    If VerificaGrupoMantenimientoUsuarios Then
        TlbMenu.Buttons.Item(6).Visible = True
        TlbMenu.Buttons.Item(6).Enabled = True
    Else
        TlbMenu.Buttons.Item(6).Visible = False
        TlbMenu.Buttons.Item(6).Enabled = False
    End If
End Sub

Private Function VerificaGrupoMantenimientoUsuarios() As Boolean
    Dim oCons As New COMDConstSistema.DCOMGeneral
    Dim sGrupoAutorizado As String
    Dim nGrupoTmp1 As String
    Dim nGrupoTmp2 As String
    Dim i As Integer
    Dim j As Integer
            
    sGrupoAutorizado = oCons.LeeConstSistema(519)
    VerificaGrupoMantenimientoUsuarios = False
    For i = 1 To Len(sGrupoAutorizado)
        If Not Mid(sGrupoAutorizado, i, 1) = "," Then
            nGrupoTmp1 = nGrupoTmp1 & Mid(sGrupoAutorizado, i, 1)
        Else
            For j = 1 To Len(gsGruposUser)
                If Not Mid(gsGruposUser, j, 1) = "," Then
                    nGrupoTmp2 = nGrupoTmp2 & Mid(gsGruposUser, j, 1)
                Else
                    If nGrupoTmp1 = nGrupoTmp2 Then
                        VerificaGrupoMantenimientoUsuarios = True
                        Exit Function
                    End If
                    nGrupoTmp2 = ""
                End If
            Next
            nGrupoTmp1 = ""
        End If
    Next
End Function
Private Sub MDIForm_Resize()
    'Posiciona el PictureBox que es el contenedor al ancho y alto que tenga el form mdi
    On Error Resume Next
    Dim ImageWidth As Single
    Dim ImageHeight As Single

    If WindowState = vbMaximized Then
        pbxFondo.Visible = False
        pbxFondo.AutoRedraw = True
        pbxFondo.Height = Me.ScaleHeight
        pbxFondo.BorderStyle = 0
    
        ImageWidth = pbxFondo.Picture.Width * 0.566893424036281
        ImageHeight = pbxFondo.Picture.Height * 0.566893424036281
    
        pbxFondo.PaintPicture pbxFondo.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, ImageWidth, ImageHeight
        Set Me.Picture = pbxFondo.Image
        pbxFondo.Refresh
    Else
        pbxFondo.BorderStyle = 0
        pbxFondo.Visible = False
        pbxFondo.AutoRedraw = True
        pbxFondo.PaintPicture pbxFondo.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, ImageWidth, ImageHeight
        pbxFondo.Move 0, 20, Me.Width, Me.Height
        Set Me.Picture = pbxFondo.Image
        pbxFondo.Refresh
    End If
End Sub
Private Sub pbxFondo_Resize()
    Dim Pos_x As Single
    Dim Pos_y As Single
  
    Pos_x = (pbxFondo.Width - Image1.Width) / 2
    Pos_y = (pbxFondo.Height - Image1.Height) / 2
    
    'Posiciona el control Image en el centro del Picturebox
    Image1.Move Pos_x, Pos_y - 680
End Sub
'<-***** Fin LUCV20190323

