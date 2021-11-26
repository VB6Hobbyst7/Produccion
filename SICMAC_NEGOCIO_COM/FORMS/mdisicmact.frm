VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDISicmact 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema del Negocio"
   ClientHeight    =   10710
   ClientLeft      =   1500
   ClientTop       =   630
   ClientWidth     =   18360
   Icon            =   "mdisicmact.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "mdisicmact.frx":030A
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar TlbMenu 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   18360
      _ExtentX        =   32385
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
            Object.ToolTipText     =   "Posicion Cliente"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdMantenimiento"
            Object.ToolTipText     =   "Mantenimiento Permisos"
            ImageIndex      =   13
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pbxFondo 
      Align           =   1  'Align Top
      Height          =   9735
      Left            =   0
      Picture         =   "mdisicmact.frx":32ECD
      ScaleHeight     =   9675
      ScaleWidth      =   18300
      TabIndex        =   2
      Top             =   600
      Width           =   18360
      Begin VB.Image Image1 
         Height          =   11520
         Left            =   0
         Picture         =   "mdisicmact.frx":B8BEE
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   21600
      End
   End
   Begin SICMACT.Usuario Usuario 
      Left            =   480
      Top             =   720
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   720
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
            Picture         =   "mdisicmact.frx":EB7B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":EBACB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":EBDE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":EC0FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":EC419
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":EC733
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":EC8C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":ECBDF
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":EDC31
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":EEC83
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":EFCD5
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":F0D27
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":F1D79
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   0
      Top             =   720
   End
   Begin MSComctlLib.StatusBar SBBarra 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   10380
      Width           =   18360
      _ExtentX        =   32385
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
            TextSave        =   "27/09/2021"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   "02:25 PM"
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
   Begin VB.Menu M0200000000 
      Caption         =   "&Captaciones"
      Index           =   0
      Begin VB.Menu M0201000000 
         Caption         =   "&Parámetros"
         Index           =   0
         Begin VB.Menu M0201010000 
            Caption         =   "&Parámetros Captaciones"
            Index           =   0
            Begin VB.Menu M0201010100 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0201010100 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0201010000 
            Caption         =   "Comisiones Diversas Ahorros"
            Index           =   1
            Begin VB.Menu M0201010200 
               Caption         =   "Registro"
               Index           =   0
            End
            Begin VB.Menu M0201010200 
               Caption         =   "Consulta"
               Index           =   1
            End
            Begin VB.Menu M0201010200 
               Caption         =   "Mantenimiento"
               Index           =   2
            End
         End
         Begin VB.Menu M0201010000 
            Caption         =   "Comisiones Transf. Banco"
            Index           =   2
            Begin VB.Menu M0201010300 
               Caption         =   "Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0201010300 
               Caption         =   "Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0201010000 
            Caption         =   "Tarifario Giros"
            Index           =   3
            Begin VB.Menu M0201010400 
               Caption         =   "Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0201010400 
               Caption         =   "Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0201010000 
            Caption         =   "Niveles aprobación C/V ME"
            Index           =   4
            Begin VB.Menu M0201010500 
               Caption         =   "Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0201010500 
               Caption         =   "Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0201010000 
            Caption         =   "Configuración de Productos"
            Index           =   5
            Begin VB.Menu M0201010600 
               Caption         =   "Ahorros"
               Index           =   0
               Begin VB.Menu M0201010601 
                  Caption         =   "Mantenimiento"
                  Index           =   0
               End
               Begin VB.Menu M0201010601 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu M0201010600 
               Caption         =   "Plazo Fijo"
               Index           =   1
               Begin VB.Menu M0201010602 
                  Caption         =   "Mantenimiento"
                  Index           =   0
               End
               Begin VB.Menu M0201010602 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu M0201010600 
               Caption         =   "CTS"
               Index           =   2
               Begin VB.Menu M0201010603 
                  Caption         =   "Mantenimiento"
                  Index           =   0
               End
               Begin VB.Menu M0201010603 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
         End
         Begin VB.Menu M0201010000 
            Caption         =   "Grupos de Agencias"
            Index           =   6
            Begin VB.Menu M0201010700 
               Caption         =   "Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0201010700 
               Caption         =   "Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0201010000 
            Caption         =   "Prog. Tarifario"
            Index           =   7
            Begin VB.Menu M0201010800 
               Caption         =   "Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0201010800 
               Caption         =   "Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0201010000 
            Caption         =   "Comisiones"
            Index           =   8
            Begin VB.Menu M0201010900 
               Caption         =   "Ahorros"
               Index           =   0
               Begin VB.Menu M0201010901 
                  Caption         =   "Mantenimiento"
                  Index           =   0
               End
               Begin VB.Menu M0201010901 
                  Caption         =   "Consulta"
                  Index           =   1
               End
               Begin VB.Menu M0201010901 
                  Caption         =   "Conf. Operaciones Interplaza"
                  Index           =   2
               End
            End
         End
         Begin VB.Menu M0201010000 
            Caption         =   "Informe de Cambio de Tarifario"
            Index           =   9
            Begin VB.Menu M0201011001 
               Caption         =   "Listar Pendientes"
               Index           =   0
            End
            Begin VB.Menu M0201011001 
               Caption         =   "Registrar Entrega de Carta"
               Index           =   1
            End
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Tasas de Interés"
         Index           =   1
         Begin VB.Menu M0201020000 
            Caption         =   "&Ahorros"
            Index           =   0
            Begin VB.Menu M0201020100 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0201020100 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0201020000 
            Caption         =   "&Plazo Fijo"
            Index           =   1
            Begin VB.Menu M0201020200 
               Caption         =   "&Mantenimiento"
               Index           =   0
               Begin VB.Menu M0201020201 
                  Caption         =   "&Persona Natural"
                  Index           =   0
               End
               Begin VB.Menu M0201020201 
                  Caption         =   "&Persona Juridica"
                  Index           =   1
               End
            End
            Begin VB.Menu M0201020200 
               Caption         =   "&Consulta"
               Index           =   1
               Begin VB.Menu M0201020202 
                  Caption         =   "&Persona Natural"
                  Index           =   0
               End
               Begin VB.Menu M0201020202 
                  Caption         =   "&Persona Juridica"
                  Index           =   1
               End
            End
            Begin VB.Menu M0201020200 
               Caption         =   "Cambio de &Tasa"
               Index           =   2
            End
            Begin VB.Menu M0201020200 
               Caption         =   " A Tasa Pactada"
               Index           =   3
            End
         End
         Begin VB.Menu M0201020000 
            Caption         =   "&CTS"
            Index           =   2
            Begin VB.Menu M0201020300 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0201020300 
               Caption         =   "&Consulta"
               Index           =   1
            End
            Begin VB.Menu M0201020300 
               Caption         =   "Cambio de Tasa Lote"
               Index           =   2
            End
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Mantenimiento"
         Index           =   2
         Begin VB.Menu M0201030000 
            Caption         =   "&Ahorros"
            Index           =   0
         End
         Begin VB.Menu M0201030000 
            Caption         =   "&Plazo Fijo"
            Index           =   1
         End
         Begin VB.Menu M0201030000 
            Caption         =   "&CTS"
            Index           =   2
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Bloqueo/Desbloqueo"
         Index           =   3
         Begin VB.Menu M0201040000 
            Caption         =   "&Ahorros"
            Index           =   0
         End
         Begin VB.Menu M0201040000 
            Caption         =   "&Plazo Fijo"
            Index           =   1
         End
         Begin VB.Menu M0201040000 
            Caption         =   "&CTS"
            Index           =   2
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Simulación Plazo Fijo"
         Index           =   4
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Tarjeta Magnética"
         Index           =   5
         Begin VB.Menu M0201050000 
            Caption         =   "&Registro"
            Index           =   0
         End
         Begin VB.Menu M0201050000 
            Caption         =   "&Bloqueo/Desbloqueo"
            Index           =   1
         End
         Begin VB.Menu M0201050000 
            Caption         =   "&Cambio de Clave"
            Index           =   2
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Bene&ficiarios"
         Index           =   6
         Begin VB.Menu M0201060000 
            Caption         =   "&Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu M0201060000 
            Caption         =   "&Consulta "
            Index           =   1
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Orden Pago"
         Index           =   7
         Begin VB.Menu M0201070000 
            Caption         =   "&Generación y Emisión"
            Index           =   0
            Begin VB.Menu M0201070100 
               Caption         =   "&Solicitud de Emisión"
               Index           =   0
            End
            Begin VB.Menu M0201070100 
               Caption         =   "&Emite OP"
               Index           =   1
            End
         End
         Begin VB.Menu M0201070000 
            Caption         =   "&Certificación"
            Index           =   1
         End
         Begin VB.Menu M0201070000 
            Caption         =   "&Anulación"
            Index           =   2
         End
         Begin VB.Menu M0201070000 
            Caption         =   "Con&sulta"
            Index           =   3
         End
         Begin VB.Menu M0201070000 
            Caption         =   "Registro Talonario"
            Index           =   4
         End
         Begin VB.Menu M0201070000 
            Caption         =   "Desbloqueo de Cliente"
            Index           =   6
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Tasas de Interés Campaña"
         Index           =   8
         Begin VB.Menu M0201080000 
            Caption         =   "Registro Campañas"
            Index           =   0
         End
         Begin VB.Menu M0201080000 
            Caption         =   "Ahorros"
            Index           =   1
            Begin VB.Menu M0201080100 
               Caption         =   "Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0201080100 
               Caption         =   "Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0201080000 
            Caption         =   "Plazo Fijo"
            Index           =   2
            Begin VB.Menu M0201080200 
               Caption         =   "Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0201080200 
               Caption         =   "Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0201080000 
            Caption         =   "CTS"
            Index           =   3
            Begin VB.Menu M0201080300 
               Caption         =   "Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0201080300 
               Caption         =   "Consulta"
               Index           =   1
            End
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Reportes"
         Index           =   9
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Generar ACL Ahorros"
         Index           =   10
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Generar ACL Colocaciones"
         Index           =   11
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Parámetros por Personería"
         Index           =   12
         Begin VB.Menu M0201090000 
            Caption         =   "&Ahorros"
            Index           =   0
         End
         Begin VB.Menu M0201090000 
            Caption         =   "&Plazo Fijo"
            Index           =   1
         End
         Begin VB.Menu M0201090000 
            Caption         =   "&CTS"
            Index           =   2
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Autorización"
         Index           =   13
         Begin VB.Menu M0201100000 
            Caption         =   "Niveles"
            Index           =   0
         End
         Begin VB.Menu M0201100000 
            Caption         =   "Niveles-Grupos"
            Index           =   1
         End
         Begin VB.Menu M0201100000 
            Caption         =   "Aprobacion / Rechazo"
            Index           =   2
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Tasas Preferenciales"
         Index           =   14
         Begin VB.Menu M0201110000 
            Caption         =   "&Solicitud"
            Index           =   0
         End
         Begin VB.Menu M0201110000 
            Caption         =   "&Aprobación/Rechazo"
            Index           =   1
         End
         Begin VB.Menu M0201110000 
            Caption         =   "&Extorno"
            Index           =   2
         End
         Begin VB.Menu M0201110000 
            Caption         =   "&Consulta"
            Index           =   3
         End
         Begin VB.Menu M0201110000 
            Caption         =   "A&dministracion de Niveles"
            Index           =   4
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Exoneración de ITF"
         Index           =   15
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Exoneraciones Descuento Inactivas"
         Index           =   16
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Bloqueo/Desbloqueo Parciales"
         Index           =   20
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&DesBloqueo de Plazo Fijo"
         Index           =   21
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Aprobar TC Especial"
         Index           =   22
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Campañas y Premios"
         Index           =   23
         Begin VB.Menu M0201120000 
            Caption         =   "Registro Campaña/Premio"
            Index           =   0
         End
         Begin VB.Menu M0201120000 
            Caption         =   "Asignación de Premio"
            Index           =   1
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Ahorro Diario"
         Index           =   24
         Begin VB.Menu M0201130000 
            Caption         =   "Renovación de Plazo"
            Index           =   0
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Reasignacion de Cuentas PF y CTS"
         Index           =   25
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Servicios"
         Index           =   26
         Begin VB.Menu M0201140000 
            Caption         =   "Gestion Convenios"
            Index           =   0
         End
         Begin VB.Menu M0201140000 
            Caption         =   "Procesar Archivo"
            Index           =   1
         End
         Begin VB.Menu M0201140000 
            Caption         =   "Reporte Pago Realizados"
            Index           =   2
         End
         Begin VB.Menu M0201140000 
            Caption         =   "Recaudo"
            Index           =   3
            Begin VB.Menu M0201140100 
               Caption         =   "Registro de Convenios"
               Index           =   0
            End
            Begin VB.Menu M0201140100 
               Caption         =   "Mantenimiento de Convenios"
               Index           =   1
            End
            Begin VB.Menu M0201140100 
               Caption         =   "Consulta de Convenios"
               Index           =   2
            End
         End
         Begin VB.Menu M0201140000 
            Caption         =   "Carga de Archivo de Trama"
            Index           =   4
         End
         Begin VB.Menu M0201140000 
            Caption         =   "Generar Trama de Retorno"
            Index           =   5
         End
         Begin VB.Menu M0201140000 
            Caption         =   "Convenio SP"
            Index           =   6
            Begin VB.Menu M0201140200 
               Caption         =   "Registro de Convenios SP"
               Index           =   0
            End
            Begin VB.Menu M0201140200 
               Caption         =   "Mantenimiento de Convenios SP"
               Index           =   1
            End
            Begin VB.Menu M0201140200 
               Caption         =   "Carga de Archivo de trama servicio"
               Index           =   2
            End
            Begin VB.Menu M0201140200 
               Caption         =   "Baja de trama servicio"
               Index           =   3
            End
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Mant.Sueldos Clientes CTS"
         Index           =   27
         Begin VB.Menu M0201150000 
            Caption         =   "Registro Manual"
            Index           =   0
         End
         Begin VB.Menu M0201150000 
            Caption         =   "Registro en Lote"
            Index           =   1
         End
         Begin VB.Menu M0201150000 
            Caption         =   "Cambio de Estado CTS"
            Index           =   2
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Tarifarios"
         Enabled         =   0   'False
         Index           =   28
         Begin VB.Menu M0201160000 
            Caption         =   "Nro. de Operaciones Max. de Retiro"
            Index           =   0
         End
         Begin VB.Menu M0201160000 
            Caption         =   "Consulta Nro. de Operaciones Max. de Retiro"
            Index           =   1
         End
         Begin VB.Menu M0201160000 
            Caption         =   "Operaciones en Otras Agencia"
            Index           =   2
         End
         Begin VB.Menu M0201160000 
            Caption         =   "Consulta Operaciones en Otras Agencia"
            Index           =   3
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Parametros PIT"
         Index           =   29
         Begin VB.Menu M0201170000 
            Caption         =   "Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu M0201170000 
            Caption         =   "Consulta"
            Index           =   1
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Monitoreo de Depósitos"
         Index           =   30
         Begin VB.Menu M0201180000 
            Caption         =   "Concentración de Clientes"
            Index           =   0
            Begin VB.Menu M0201180100 
               Caption         =   "Definir Parametros"
               Index           =   0
            End
            Begin VB.Menu M0201180100 
               Caption         =   "Autorización de Depósito"
               Index           =   1
            End
         End
         Begin VB.Menu M0201190000 
            Caption         =   "Comportamiento de depósitos"
            Index           =   0
            Begin VB.Menu M0201180200 
               Caption         =   "Definir Parámetros"
               Index           =   0
            End
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Requerimientos SUNAT"
         Index           =   31
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Aprobar TC Espec"
         Index           =   32
      End
   End
   Begin VB.Menu M0300000000 
      Caption         =   "Colocacio&nes"
      Index           =   0
      Begin VB.Menu M0301000000 
         Caption         =   "Cartas &Fianza"
         Index           =   0
         Begin VB.Menu M0301010000 
            Caption         =   "&Solicitud Carta Fianza"
            Index           =   0
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Gra&var Garantia"
            Index           =   1
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Sugerencia de &Analista"
            Index           =   2
         End
         Begin VB.Menu M0301010000 
            Caption         =   "&Resolver Carta Fianza"
            Index           =   3
            Begin VB.Menu M0301010100 
               Caption         =   "&Aprobar Carta Fianza"
               Index           =   0
            End
            Begin VB.Menu M0301010100 
               Caption         =   "&Rechazar Carta Fianza"
               Index           =   1
            End
            Begin VB.Menu M0301010100 
               Caption         =   "Retirar Carta Fianza Aprobada"
               Index           =   2
            End
            Begin VB.Menu M0301010100 
               Caption         =   "&Devolver Carta Fianza Emitida"
               Index           =   3
            End
            Begin VB.Menu M0301010100 
               Caption         =   "&Cancelar Carta Fianza Emitida"
               Index           =   4
            End
            Begin VB.Menu M0301010100 
               Caption         =   "Ex&torno Carta Fianza Aprobada "
               Index           =   5
            End
            Begin VB.Menu M0301010100 
               Caption         =   "Modificar Carta Fianza"
               Index           =   6
            End
            Begin VB.Menu M0301010100 
               Caption         =   "Editar Modalidad CF Emitida"
               Index           =   7
            End
         End
         Begin VB.Menu M0301010000 
            Caption         =   "&Emitir Carta Fianza"
            Index           =   4
         End
         Begin VB.Menu M0301010000 
            Caption         =   "&Honrar Carta Fianza"
            Index           =   5
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Consultas"
            Index           =   6
            Begin VB.Menu M0301010200 
               Caption         =   "Historial de Carta Fianza"
               Index           =   0
            End
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Reportes"
            Index           =   7
            Begin VB.Menu M0301010300 
               Caption         =   "Reportes de Carta Fianza"
               Index           =   0
            End
         End
         Begin VB.Menu M0301010000 
            Caption         =   "&Manteminiento Tarifario"
            Index           =   8
         End
         Begin VB.Menu M0301010000 
            Caption         =   "&Relacionar con Credito"
            Index           =   9
         End
         Begin VB.Menu M0301010000 
            Caption         =   "&Renovacion"
            Index           =   10
            Begin VB.Menu M0301010600 
               Caption         =   "Autorización Pago Por Renovación"
               Index           =   0
            End
            Begin VB.Menu M0301010600 
               Caption         =   "Extorno de Autorización de Renovación"
               Index           =   1
            End
            Begin VB.Menu M0301010600 
               Caption         =   "Renovación Carta Fianza"
               Index           =   2
            End
            Begin VB.Menu M0301010600 
               Caption         =   "Extorno de Renovación"
               Index           =   3
            End
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Dar de Baja"
            Index           =   11
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Mantenimiento de Modalidad"
            Index           =   12
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Hoja de Aprobación"
            Index           =   13
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "&Creditos"
         Index           =   1
         Begin VB.Menu M0301020000 
            Caption         =   "&Definiciones"
            Index           =   0
            Begin VB.Menu M0301020100 
               Caption         =   "&Parametros de Control"
               Index           =   0
               Begin VB.Menu M0301020101 
                  Caption         =   "&Mantenimiento"
                  Index           =   0
               End
               Begin VB.Menu M0301020101 
                  Caption         =   "&Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu M0301020100 
               Caption         =   "&Lineas de Credito"
               Enabled         =   0   'False
               Index           =   1
               Visible         =   0   'False
               Begin VB.Menu M0301020102 
                  Caption         =   "&Registro"
                  Index           =   0
               End
               Begin VB.Menu M0301020102 
                  Caption         =   "&Mantenimiento"
                  Index           =   1
               End
               Begin VB.Menu M0301020102 
                  Caption         =   "&Consulta"
                  Index           =   2
               End
            End
            Begin VB.Menu M0301020100 
               Caption         =   "&Niveles de Aprobacion"
               Index           =   2
               Begin VB.Menu M0301020103 
                  Caption         =   "Grupo de Aprobación Créditos"
                  Index           =   0
               End
               Begin VB.Menu M0301020103 
                  Caption         =   "Niveles de Aprobación"
                  Index           =   1
               End
               Begin VB.Menu M0301020103 
                  Caption         =   "Parámetros de Aprobación"
                  Index           =   2
               End
               Begin VB.Menu M0301020103 
                  Caption         =   "Delegación de Nivel de Arpobación"
                  Index           =   3
               End
               Begin VB.Menu M0301020103 
                  Caption         =   "Configurar Cargos para aprobación de tasa"
                  Index           =   4
               End
            End
            Begin VB.Menu M0301020100 
               Caption         =   "&Gastos"
               Index           =   3
               Begin VB.Menu M0301020104 
                  Caption         =   "&Mantenimiento"
                  Index           =   0
               End
               Begin VB.Menu M0301020104 
                  Caption         =   "&Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu M0301020100 
               Caption         =   "Campañas"
               Index           =   5
               Begin VB.Menu M0301020105 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu M0301020105 
                  Caption         =   "Mantenimiento"
                  Index           =   1
               End
               Begin VB.Menu M0301020105 
                  Caption         =   "Consultas"
                  Index           =   2
               End
               Begin VB.Menu M0301020105 
                  Caption         =   "Casa Comercial"
                  Index           =   3
               End
            End
            Begin VB.Menu M0301020100 
               Caption         =   "Parámetros de Evaluación"
               Index           =   6
               Begin VB.Menu M0301020106 
                  Caption         =   "Tipo de Evaluación"
                  Enabled         =   0   'False
                  Index           =   0
                  Visible         =   0   'False
               End
               Begin VB.Menu M0301020106 
                  Caption         =   "Indicadores de Evaluación"
                  Enabled         =   0   'False
                  Index           =   1
                  Visible         =   0   'False
               End
               Begin VB.Menu M0301020106 
                  Caption         =   "Especialización de Créditos"
                  Enabled         =   0   'False
                  Index           =   2
                  Visible         =   0   'False
               End
               Begin VB.Menu M0301020106 
                  Caption         =   "Configuración de Parámetros de Evaluación"
                  Index           =   3
               End
               Begin VB.Menu M0301020106 
                  Caption         =   "Configuración de Ratios por Tipo de Productos Crediticios"
                  Index           =   4
               End
            End
            Begin VB.Menu M0301020100 
               Caption         =   "Niveles de Autorización"
               Index           =   7
               Begin VB.Menu M0301020107 
                  Caption         =   "Tipos de Autorización"
                  Index           =   0
               End
               Begin VB.Menu M0301020107 
                  Caption         =   "Niveles de Autorización"
                  Index           =   1
               End
               Begin VB.Menu M0301020107 
                  Caption         =   "Configurar Autorización"
                  Index           =   2
               End
            End
            Begin VB.Menu M0301020100 
               Caption         =   "Periodo maximo por destino"
               Index           =   8
               Begin VB.Menu M0301020108 
                  Caption         =   "Mantenimiento"
                  Index           =   0
               End
               Begin VB.Menu M0301020108 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu M0301020100 
               Caption         =   "Administración de Productos"
               Index           =   9
               Begin VB.Menu M0301020109 
                  Caption         =   "SubProducto Cuota Balón"
                  Index           =   0
               End
               Begin VB.Menu M0301020109 
                  Caption         =   "Configuración de Tarifarios"
                  Index           =   1
               End
               Begin VB.Menu M0301020109 
                  Caption         =   "Configuración de Catálogo de Productos"
                  Index           =   2
               End
            End
            Begin VB.Menu M0301020100 
               Caption         =   "Clientes Preferenciales"
               Index           =   10
            End
            Begin VB.Menu M0301020100 
               Caption         =   "Límites por Sector Econ"
               Index           =   11
               Begin VB.Menu M0301020111 
                  Caption         =   "Mantenimiento"
                  Index           =   0
               End
               Begin VB.Menu M0301020111 
                  Caption         =   "Consulta"
                  Index           =   1
               End
               Begin VB.Menu M0301020111 
                  Caption         =   "Configuración de límites de alertas tempranas"
                  Index           =   2
               End
            End
            Begin VB.Menu M0301020100 
               Caption         =   "Comisiones Diversas Créditos"
               Index           =   12
               Begin VB.Menu M0301020110 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu M0301020110 
                  Caption         =   "Mantenimiento"
                  Index           =   1
               End
               Begin VB.Menu M0301020110 
                  Caption         =   "Consulta"
                  Index           =   2
               End
            End
            Begin VB.Menu M0301020100 
               Caption         =   "Límites por Zona Geog."
               Index           =   13
            End
            Begin VB.Menu M0301020100 
               Caption         =   "Límites por Tipo de Producto."
               Index           =   14
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Solicitud Credito"
            Index           =   1
            Begin VB.Menu M0301020200 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M0301020200 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Evaluación"
            Index           =   2
            Begin VB.Menu M0301024500 
               Caption         =   "Registro"
               Index           =   0
            End
            Begin VB.Menu M0301024500 
               Caption         =   "Matenimiento"
               Index           =   1
            End
            Begin VB.Menu M0301024500 
               Caption         =   "Consulta"
               Index           =   2
            End
            Begin VB.Menu M0301024500 
               Caption         =   "Fuente de Ingreso"
               Index           =   3
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Relaciones de Credito"
            Index           =   3
            Begin VB.Menu M0301020300 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0301020300 
               Caption         =   "Reasignacion de &Cartera en Lote"
               Index           =   1
            End
            Begin VB.Menu M0301020300 
               Caption         =   "Con&sulta"
               Index           =   2
            End
            Begin VB.Menu M0301020300 
               Caption         =   "Confirmación de &Reasignación de Cartera"
               Index           =   3
            End
            Begin VB.Menu M0301020300 
               Caption         =   "Asignacion de Agencia - JEFE DE NEGOCIO TERRITORIAL"
               Index           =   4
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Garantias"
            Index           =   4
            Begin VB.Menu M0301020400 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M0301020400 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M0301020400 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu M0301020400 
               Caption         =   "Registro de Cobertura"
               Index           =   3
            End
            Begin VB.Menu M0301020400 
               Caption         =   "Configurar"
               Index           =   5
            End
            Begin VB.Menu M0301020400 
               Caption         =   "Bloqueo/Desbloqueo Legal"
               Index           =   10
            End
            Begin VB.Menu M0301020400 
               Caption         =   "Verifica Legal"
               Index           =   11
            End
            Begin VB.Menu M0301020400 
               Caption         =   "Cambio de Garantía"
               Enabled         =   0   'False
               Index           =   12
            End
            Begin VB.Menu M0301020400 
               Caption         =   "Cartas de Vencimiento de Tasación"
               Enabled         =   0   'False
               Index           =   13
            End
            Begin VB.Menu M0301020400 
               Caption         =   "Configurar Coberturas x Producto"
               Index           =   14
            End
            Begin VB.Menu M0301020400 
               Caption         =   "Exoneración Ratio de Cobertura"
               Index           =   15
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Sugerencia"
            Index           =   5
            Begin VB.Menu M0301020500 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M0301020500 
               Caption         =   "&Consulta"
               Index           =   1
            End
            Begin VB.Menu M0301020500 
               Caption         =   "Desbloqueo Sobreendeudamiento"
               Index           =   2
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Resolver Creditos"
            Index           =   6
            Begin VB.Menu M0301020600 
               Caption         =   "&Aprobacion"
               Index           =   0
            End
            Begin VB.Menu M0301020600 
               Caption         =   "&Rechazo"
               Index           =   1
            End
            Begin VB.Menu M0301020600 
               Caption         =   "A&nulacion Creditos Aprobados"
               Index           =   2
            End
            Begin VB.Menu M0301020600 
               Caption         =   "Extornado"
               Index           =   3
            End
            Begin VB.Menu M0301020600 
               Caption         =   "Desbloqueo de Credito"
               Index           =   4
            End
            Begin VB.Menu M0301020600 
               Caption         =   "Rechazar solicitud"
               Index           =   5
            End
            Begin VB.Menu M0301020600 
               Caption         =   "Rechazar Sugerencia"
               Index           =   6
            End
            Begin VB.Menu M0301020600 
               Caption         =   "Aprobacion Por Niveles"
               Index           =   7
            End
            Begin VB.Menu M0301020600 
               Caption         =   "Historial de Niveles de Aprobacion"
               Index           =   8
            End
            Begin VB.Menu M0301020600 
               Caption         =   "Aprobar solicitud de tasa en aprobación"
               Index           =   9
            End
            Begin VB.Menu M0301020600 
               Caption         =   "Extorno Aprobación Por Niveles"
               Index           =   10
            End
            Begin VB.Menu M0301020600 
               Caption         =   "Historial de Autorizaciones"
               Index           =   11
            End
            Begin VB.Menu M0301020600 
               Caption         =   "Autorización por Niveles"
               Index           =   12
            End
            Begin VB.Menu M0301020600 
               Caption         =   "VB por deuda potencial"
               Index           =   13
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Reprogramacion Credito"
            Index           =   7
            Begin VB.Menu M0301020700 
               Caption         =   "Solicitud"
               Index           =   0
            End
            Begin VB.Menu M0301020700 
               Caption         =   "Propuesta"
               Index           =   1
            End
            Begin VB.Menu M0301020700 
               Caption         =   "Autorización"
               Index           =   2
            End
            Begin VB.Menu M0301020700 
               Caption         =   "Aprobación"
               Index           =   3
            End
            Begin VB.Menu M0301020700 
               Caption         =   "Rechazo"
               Index           =   4
            End
            Begin VB.Menu M0301020700 
               Caption         =   "Repr&ogramacion"
               Index           =   5
            End
            Begin VB.Menu M0301020700 
               Caption         =   "Registro V°B° Adm Creditos"
               Index           =   6
            End
            Begin VB.Menu M0301020700 
               Caption         =   "Reprogramacion en &Lote"
               Index           =   7
            End
            Begin VB.Menu M0301020700 
               Caption         =   "Ampliación de Plazo Créditos Convenio"
               Index           =   8
            End
            Begin VB.Menu M0301020700 
               Caption         =   "Extorno de Reprogramación"
               Index           =   9
            End
            Begin VB.Menu M0301020700 
               Caption         =   "Garantía Covid"
               Index           =   10
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Refinanciacion"
            Index           =   8
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Actualizacion de &Metodos de Liquidacion"
            Index           =   9
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Perdonar Mora"
            Index           =   10
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Gastos"
            Index           =   11
            Begin VB.Menu M0301020800 
               Caption         =   "&Administracion de Gastos en &Lote"
               Index           =   0
            End
            Begin VB.Menu M0301020800 
               Caption         =   "Mantenimiento de Penalidad de Cancelacion"
               Index           =   1
            End
            Begin VB.Menu M0301020800 
               Caption         =   "Administracion de Gastos"
               Index           =   2
            End
            Begin VB.Menu M0301020800 
               Caption         =   "Asignacion de Gastos en Lote"
               Index           =   3
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Reasignar &Institucion"
            Index           =   12
            Visible         =   0   'False
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Transferencia a Recuperaciones"
            Index           =   13
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Analista"
            Index           =   14
            Begin VB.Menu M0301020900 
               Caption         =   "&Nota de Analista"
               Index           =   0
            End
            Begin VB.Menu M0301020900 
               Caption         =   "&Metas de Analista"
               Index           =   1
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Si&mulaciones"
            Index           =   15
            Begin VB.Menu M0301021000 
               Caption         =   "Calendario de &Pagos"
               Index           =   0
            End
            Begin VB.Menu M0301021000 
               Caption         =   "Calendario de &Desembolsos Parciales"
               Index           =   1
            End
            Begin VB.Menu M0301021000 
               Caption         =   "Calendario de Cuota &Libre"
               Index           =   2
            End
            Begin VB.Menu M0301021000 
               Caption         =   "Simulador de Pagos"
               Index           =   3
            End
            Begin VB.Menu M0301021000 
               Caption         =   "Simulador Créditos Rapiflash"
               Index           =   4
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Cons&ultas"
            Index           =   16
            Begin VB.Menu M0301021100 
               Caption         =   "&Historial de Credito"
               Index           =   0
            End
            Begin VB.Menu M0301021100 
               Caption         =   "Historial de &Calendario"
               Index           =   1
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Reportes"
            Index           =   17
            Begin VB.Menu M0301021200 
               Caption         =   "&Duplicados"
               Index           =   0
            End
            Begin VB.Menu M0301021200 
               Caption         =   "Listados de Creditos"
               Index           =   1
            End
            Begin VB.Menu M0301021200 
               Caption         =   "Creditos Vinculados Titulares"
               Index           =   2
            End
            Begin VB.Menu M0301021200 
               Caption         =   "Creditos Vinculados Titulares y Garantes"
               Index           =   3
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Registro de Dacion de Pago"
            Index           =   18
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Cargo Automatico"
            Index           =   19
            Begin VB.Menu M0301021500 
               Caption         =   "Asignar Cargo Automatico"
               Index           =   1
            End
            Begin VB.Menu M0301021500 
               Caption         =   "Mantenimiento"
               Index           =   2
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Modificar Codigos Modulares"
            Index           =   20
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Asignar Cuota Comodin"
            Index           =   21
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Administrar PrePagos"
            Index           =   22
            Begin VB.Menu M0301021600 
               Caption         =   "PrePagos Hipotecarios"
               Index           =   1
            End
            Begin VB.Menu M0301021600 
               Caption         =   "PrePagos Normales"
               Index           =   2
            End
            Begin VB.Menu M0301021600 
               Caption         =   "Exoneración de Comisión de Precancelación"
               Index           =   3
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Calendario de Desembolsos"
            Index           =   24
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Sustitución de Deudor"
            Index           =   25
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Ingreso Devolucion de Cred.Convenio"
            Index           =   27
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Autorizar Credito"
            Index           =   29
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Cambio de Linea de Credito"
            Index           =   31
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Ampliacion de Credito"
            Index           =   32
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Polizas"
            Enabled         =   0   'False
            Index           =   33
            Begin VB.Menu M0301021700 
               Caption         =   "Operaciones"
               Enabled         =   0   'False
               Index           =   0
               Begin VB.Menu M0301021701 
                  Caption         =   "Registro de Póliza Pendiente"
                  Enabled         =   0   'False
                  Index           =   0
               End
               Begin VB.Menu M0301021701 
                  Caption         =   "Registro de Póliza Manual"
                  Enabled         =   0   'False
                  Index           =   1
               End
               Begin VB.Menu M0301021701 
                  Caption         =   "Registro de Póliza en Crédito Vigente"
                  Enabled         =   0   'False
                  Index           =   2
               End
               Begin VB.Menu M0301021701 
                  Caption         =   "Renovación Póliza Externa"
                  Enabled         =   0   'False
                  Index           =   3
               End
               Begin VB.Menu M0301021701 
                  Caption         =   "Sustitución de Póliza"
                  Enabled         =   0   'False
                  Index           =   4
               End
               Begin VB.Menu M0301021701 
                  Caption         =   "Relación Póliza - Garantía"
                  Enabled         =   0   'False
                  Index           =   5
               End
            End
            Begin VB.Menu M0301021700 
               Caption         =   "Consulta"
               Enabled         =   0   'False
               Index           =   1
            End
            Begin VB.Menu M0301021700 
               Caption         =   "Cotizador de Seguro"
               Enabled         =   0   'False
               Index           =   2
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Corresponsalia"
            Index           =   34
            Begin VB.Menu M0301021800 
               Caption         =   "Pagos en Banco de la Nación"
               Index           =   0
            End
            Begin VB.Menu M0301021800 
               Caption         =   "Desembolso en Banco de la Nación"
               Index           =   1
            End
            Begin VB.Menu M0301021800 
               Caption         =   "Recuperaciones en Banco de la Nación"
               Index           =   2
            End
            Begin VB.Menu M0301021800 
               Caption         =   "Reportes Pagos Automatico Corresponsalia Banco de la Nación"
               Index           =   3
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Gestión de Cobranza"
            Index           =   35
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Vinculados y Grupos Económicos"
            Index           =   36
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Comite"
            Index           =   37
            Begin VB.Menu M0301021900 
               Caption         =   "Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0301021900 
               Caption         =   "Reporte Actas Comite"
               Index           =   1
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Constancia de No Adeudo"
            Enabled         =   0   'False
            Index           =   38
            Begin VB.Menu M0301022000 
               Caption         =   "Registrar Solicitud de Constancia de No Adeudo"
               Enabled         =   0   'False
               Index           =   0
            End
            Begin VB.Menu M0301022000 
               Caption         =   "Listar Constancia de No Adeudo"
               Enabled         =   0   'False
               Index           =   1
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "BPP"
            Index           =   39
            Begin VB.Menu M0301022100 
               Caption         =   "Tipos de Cartera"
               Index           =   0
            End
            Begin VB.Menu M0301022100 
               Caption         =   "Metas Analistas"
               Index           =   1
            End
            Begin VB.Menu M0301022100 
               Caption         =   "Parametros BPP"
               Index           =   2
            End
            Begin VB.Menu M0301022100 
               Caption         =   "Configuración General"
               Index           =   3
            End
            Begin VB.Menu M0301022100 
               Caption         =   "Configuración Mensual"
               Index           =   4
            End
            Begin VB.Menu M0301022100 
               Caption         =   "Configuración Agencia"
               Index           =   5
               Begin VB.Menu M0301022106 
                  Caption         =   "Metas Mensuales"
                  Index           =   0
               End
               Begin VB.Menu M0301022106 
                  Caption         =   "Comités de Créditos"
                  Index           =   1
               End
            End
            Begin VB.Menu M0301022100 
               Caption         =   "Generación de Bono Mes"
               Index           =   6
            End
            Begin VB.Menu M0301022100 
               Caption         =   "Promotores"
               Index           =   7
               Begin VB.Menu M0301022108 
                  Caption         =   "Configuración"
                  Index           =   0
               End
               Begin VB.Menu M0301022108 
                  Caption         =   "Bonificación Mensual"
                  Index           =   1
               End
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Metas Agencias"
            Index           =   40
         End
         Begin VB.Menu M0301020000 
            Caption         =   "COFIDE"
            Index           =   41
            Begin VB.Menu M0301022200 
               Caption         =   "Registro Persona Cofide"
               Index           =   1
            End
            Begin VB.Menu M0301022200 
               Caption         =   "Registro de Vehículo"
               Index           =   2
            End
            Begin VB.Menu M0301022200 
               Caption         =   "Solicitud de Habilitación (Formato I)"
               Index           =   3
            End
            Begin VB.Menu M0301022200 
               Caption         =   "Confirmación de Habilitación (Formato III)"
               Index           =   4
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Microseguros"
            Index           =   42
            Begin VB.Menu M0301022300 
               Caption         =   "Registro"
               Index           =   1
            End
            Begin VB.Menu M0301022300 
               Caption         =   "Tramas"
               Index           =   2
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Estado"
            Index           =   43
            Begin VB.Menu M0301022400 
               Caption         =   "Asesoría Legal - Informe"
               Index           =   1
            End
            Begin VB.Menu M0301022400 
               Caption         =   "Parámetro de Revisión Sup. Cred."
               Index           =   2
            End
            Begin VB.Menu M0301022400 
               Caption         =   "Supervisión de Créditos"
               Index           =   3
            End
            Begin VB.Menu M0301022400 
               Caption         =   "Asesoría Legal - Minuta"
               Index           =   4
            End
            Begin VB.Menu M0301022400 
               Caption         =   "Seguimiento de Credito"
               Index           =   5
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Informe de Riesgos"
            Index           =   44
            Begin VB.Menu M0301022500 
               Caption         =   "Registrar Informe"
               Index           =   0
            End
            Begin VB.Menu M0301022500 
               Caption         =   "Mantenimiento de Informe"
               Index           =   1
            End
            Begin VB.Menu M0301022500 
               Caption         =   "Confirmaciòn de Comunicaciòn - Rapiflash"
               Index           =   2
            End
            Begin VB.Menu M0301022500 
               Caption         =   "Solic. Límites por Sectores"
               Index           =   3
               Begin VB.Menu M0301022501 
                  Caption         =   "Autorización/Rechazo"
                  Index           =   0
               End
               Begin VB.Menu M0301022501 
                  Caption         =   "Extorno"
                  Index           =   1
               End
            End
            Begin VB.Menu M0301022500 
               Caption         =   "Registro de Parámetros de Autorizaciones"
               Index           =   4
            End
            Begin VB.Menu M0301022500 
               Caption         =   "Registro Categoría Agencia (Mensual) "
               Index           =   5
            End
            Begin VB.Menu M0301022500 
               Caption         =   "Listado de Autorizaciones"
               Index           =   6
            End
            Begin VB.Menu M0301022500 
               Caption         =   "Configuración de Monto - Segmentación"
               Index           =   7
            End
            Begin VB.Menu M0301022500 
               Caption         =   "Configuración de Monto - Cat. Agencias"
               Index           =   8
            End
            Begin VB.Menu M0301022500 
               Caption         =   "Configuración de Códigos de Sobr."
               Index           =   9
            End
            Begin VB.Menu M0301022500 
               Caption         =   "Solic. Límite por Zona Geog."
               Index           =   10
            End
            Begin VB.Menu M0301022500 
               Caption         =   "Solic. Límite por Tipo de Producto."
               Index           =   11
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Parámetros"
            Index           =   45
            Begin VB.Menu M0301022600 
               Caption         =   "Créditos Agropecuarios"
               Index           =   0
               Begin VB.Menu M0301022601 
                  Caption         =   "Tipo de Productos Agropecuarios"
                  Index           =   0
               End
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Seguimientos de SobreEndeudados"
            Index           =   46
            Begin VB.Menu M0301022700 
               Caption         =   "Registro de Visita Analista"
               Index           =   0
            End
            Begin VB.Menu M0301022700 
               Caption         =   "Registro de Visita Jefe de Agencia"
               Index           =   1
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Traspaso de Cartera entre Agencias"
            Index           =   47
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Mantenimiento Créditos Convenio"
            Index           =   48
            Begin VB.Menu M0301022800 
               Caption         =   "Reasignar Institución"
               Index           =   0
            End
            Begin VB.Menu M0301022800 
               Caption         =   "Asignación Convenio"
               Index           =   1
            End
            Begin VB.Menu M0301022800 
               Caption         =   "Retiro Conveio"
               Index           =   2
            End
            Begin VB.Menu M0301022800 
               Caption         =   "Historial de Cambios Créditos Convenio"
               Index           =   3
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Proyecciones de Agencias"
            Index           =   49
            Begin VB.Menu M0301022900 
               Caption         =   "Seguimiento de Proyecciones Semanales"
               Index           =   0
            End
            Begin VB.Menu M0301022900 
               Caption         =   "Proyectado Vs Ejecutado por Agencia"
               Index           =   1
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "MIVIVIENDA"
            Index           =   50
            Begin VB.Menu M0301023000 
               Caption         =   "Alertas para Créditos MIVIVIENDA"
               Index           =   0
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Autorización Solicitud Ampliación"
            Index           =   51
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Ficha Sobreendeudado"
            Index           =   52
            Begin VB.Menu M0301024000 
               Caption         =   "Registro"
               Index           =   0
            End
            Begin VB.Menu M0301024000 
               Caption         =   "Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M0301024000 
               Caption         =   "Consulta"
               Index           =   2
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Asistente de Agencia"
            Index           =   53
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Registro de Solicitud del 25% del AFP"
            Index           =   54
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "&Pignoraticio"
         Index           =   2
         Begin VB.Menu M0301030000 
            Caption         =   "&Contrato"
            Index           =   0
            Begin VB.Menu M0301030100 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M0301030100 
               Caption         =   "Consulta de Descripción"
               Index           =   1
            End
            Begin VB.Menu M0301030100 
               Caption         =   "Anulación"
               Index           =   2
            End
            Begin VB.Menu M0301030100 
               Caption         =   "&Bloqueo"
               Index           =   3
            End
            Begin VB.Menu M0301030100 
               Caption         =   "Ampliación"
               Index           =   4
            End
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Rescate Joyas"
            Index           =   1
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Remate"
            Enabled         =   0   'False
            Index           =   2
            Begin VB.Menu M0301030300 
               Caption         =   "Retasación de Joyas"
               Enabled         =   0   'False
               Index           =   0
            End
            Begin VB.Menu M0301030300 
               Caption         =   "Preparacion Remate"
               Enabled         =   0   'False
               Index           =   1
            End
            Begin VB.Menu M0301030300 
               Caption         =   "Remate"
               Enabled         =   0   'False
               Index           =   2
            End
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Subasta"
            Index           =   4
            Begin VB.Menu M0301030400 
               Caption         =   "Preparacion Subasta"
               Index           =   0
            End
            Begin VB.Menu M0301030400 
               Caption         =   "Subasta"
               Index           =   1
            End
         End
         Begin VB.Menu M0301030000 
            Caption         =   "&Venta en Lote"
            Index           =   5
            Begin VB.Menu M0301030500 
               Caption         =   "Preparación de Venta en lote"
               Index           =   1
            End
            Begin VB.Menu M0301030500 
               Caption         =   "Venta en Lote"
               Index           =   2
            End
         End
         Begin VB.Menu M0301030000 
            Caption         =   "&Reportes"
            Index           =   7
            Begin VB.Menu M0301030600 
               Caption         =   "Historial de Contrato"
               Index           =   0
            End
            Begin VB.Menu M0301030600 
               Caption         =   "Contrato por Persona"
               Index           =   1
            End
            Begin VB.Menu M0301030600 
               Caption         =   "-"
               Index           =   2
            End
            Begin VB.Menu M0301030600 
               Caption         =   "Movimientos Diario"
               Index           =   3
            End
            Begin VB.Menu M0301030600 
               Caption         =   "Listados Generales"
               Index           =   4
            End
            Begin VB.Menu M0301030600 
               Caption         =   "Estadisticas"
               Index           =   5
            End
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Adjudicacion"
            Index           =   8
            Begin VB.Menu M0301030700 
               Caption         =   "Preparacion Adjudicacion"
               Index           =   0
            End
            Begin VB.Menu M0301030700 
               Caption         =   "Adjudicacion de Lotes"
               Index           =   1
            End
            Begin VB.Menu M0301030700 
               Caption         =   "Reimpresión Comprobante"
               Index           =   2
            End
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Definiciones"
            Index           =   9
            Begin VB.Menu M0301030800 
               Caption         =   "Precio Oro"
               Index           =   0
            End
            Begin VB.Menu M0301030800 
               Caption         =   "Tarifario de Cartas Notariales - Minka"
               Index           =   1
            End
            Begin VB.Menu M0301030800 
               Caption         =   "Tarifario de Cartas Notariales - Agencia"
               Enabled         =   0   'False
               Index           =   2
            End
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Retasación"
            Index           =   10
            Begin VB.Menu M0301030900 
               Caption         =   "Preparación vigentes/Diferidas/Adjudicados"
               Index           =   0
            End
            Begin VB.Menu M0301030900 
               Caption         =   "Consulta Retasaciones"
               Index           =   1
            End
            Begin VB.Menu M0301030900 
               Caption         =   "Retasación Vigentes/Diferidas/Adjudicados"
               Index           =   2
            End
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Campaña Recuperación Adjudicados"
            Enabled         =   0   'False
            Index           =   11
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Configuración"
            Index           =   12
            Begin VB.Menu M0301031000 
               Caption         =   "Parámetros de Segmentación"
               Index           =   0
            End
            Begin VB.Menu M0301031000 
               Caption         =   "Retención de tasas "
               Index           =   1
            End
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Registro de Holograma"
            Index           =   13
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Retención de Creditos Prendario"
            Index           =   14
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "&Recuperaciones"
         Enabled         =   0   'False
         Index           =   4
         Visible         =   0   'False
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
            Caption         =   "Adjudicación/Venta de Bienes"
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
      Begin VB.Menu M0301000000 
         Caption         =   "Calificación de Colocaciones"
         Index           =   5
         Begin VB.Menu M0301060000 
            Caption         =   "Evaluación"
            Index           =   0
            Begin VB.Menu M0301060100 
               Caption         =   "Evaluacion de Cartera"
               Index           =   0
            End
            Begin VB.Menu M0301060100 
               Caption         =   "Evaluacion Automatica"
               Index           =   1
            End
            Begin VB.Menu M0301060100 
               Caption         =   "Garantias Preferidas"
               Index           =   2
            End
            Begin VB.Menu M0301060100 
               Caption         =   "Reclasificacion Mes Comercial"
               Index           =   3
            End
         End
         Begin VB.Menu M0301060000 
            Caption         =   "Preparan Data"
            Index           =   1
         End
         Begin VB.Menu M0301060000 
            Caption         =   "Calificacion"
            Index           =   2
            Begin VB.Menu M0301060200 
               Caption         =   "Calificacion 11356"
               Index           =   0
            End
            Begin VB.Menu M0301060200 
               Caption         =   "Cierre de Calificacion"
               Index           =   1
            End
            Begin VB.Menu M0301060200 
               Caption         =   "Parametros"
               Index           =   2
            End
         End
         Begin VB.Menu M0301060000 
            Caption         =   "Reporte"
            Index           =   3
         End
         Begin VB.Menu M0301060000 
            Caption         =   "Consulta de Calificacion"
            Index           =   4
         End
         Begin VB.Menu M0301060000 
            Caption         =   "Datos Reportes"
            Index           =   5
            Begin VB.Menu M0301060300 
               Caption         =   "Interes Devengados"
               Index           =   1
            End
            Begin VB.Menu M0301060300 
               Caption         =   "Activos y Contingentes Ponderados por Riesgo de Crédito"
               Index           =   2
            End
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "Informe RCD"
         Index           =   6
         Begin VB.Menu M0301070000 
            Caption         =   "Parametros RCD"
            Index           =   0
         End
         Begin VB.Menu M0301070000 
            Caption         =   "Preparacion de Data"
            Index           =   1
            Begin VB.Menu M0301070100 
               Caption         =   "Datos Maestro RCD"
               Index           =   0
            End
            Begin VB.Menu M0301070100 
               Caption         =   "Persona desde MaestroRCD"
               Index           =   1
            End
         End
         Begin VB.Menu M0301070000 
            Caption         =   "Generar Datos"
            Index           =   2
            Begin VB.Menu M0301070200 
               Caption         =   "Informe RCD"
               Index           =   0
            End
            Begin VB.Menu M0301070200 
               Caption         =   "Informe IBM"
               Index           =   1
               Visible         =   0   'False
            End
            Begin VB.Menu M0301070200 
               Caption         =   "Verificar Datos"
               Index           =   2
            End
         End
         Begin VB.Menu M0301070000 
            Caption         =   "Reportes"
            Index           =   3
         End
         Begin VB.Menu M0301070000 
            Caption         =   "Validacion de Personas"
            Index           =   4
            Begin VB.Menu M0301070300 
               Caption         =   "RCDMaestro Persona"
               Index           =   0
            End
         End
         Begin VB.Menu M0301070000 
            Caption         =   "Modificar Datos RCD"
            Index           =   5
         End
         Begin VB.Menu M0301070000 
            Caption         =   "Convertir Observación"
            Index           =   6
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "Arrendmiento Financiero"
         Enabled         =   0   'False
         Index           =   7
         Visible         =   0   'False
         Begin VB.Menu M0301080000 
            Caption         =   "Solicitud Arrendamiento Financiero"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu M0301080000 
            Caption         =   "Sugerencia Arrendamiento Financiero"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu M0301080000 
            Caption         =   "Aprobación Arrendamiento Financiero"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu M0301080000 
            Caption         =   "Desembolso Arrendamiento Financiero"
            Enabled         =   0   'False
            Index           =   3
         End
         Begin VB.Menu M0301080000 
            Caption         =   "Cuota Inicial Arrendamiento Financiero"
            Enabled         =   0   'False
            Index           =   4
         End
         Begin VB.Menu M0301080000 
            Caption         =   "Usuario Leasing"
            Enabled         =   0   'False
            Index           =   5
            Visible         =   0   'False
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "Hoja de Ruta"
         Index           =   8
         Begin VB.Menu M0301090000 
            Caption         =   "Registro"
            Index           =   0
         End
         Begin VB.Menu M0301090000 
            Caption         =   "Mantenimiento"
            Index           =   1
         End
         Begin VB.Menu M0301090000 
            Caption         =   "Mantenimiento Coordinador"
            Index           =   2
         End
         Begin VB.Menu M0301090000 
            Caption         =   "Resultado de Visitas"
            Index           =   3
         End
         Begin VB.Menu M0301090000 
            Caption         =   "Consulta Hoja de Ruta"
            Index           =   4
         End
         Begin VB.Menu M0301090000 
            Caption         =   "Resultado de Visita"
            Index           =   5
         End
         Begin VB.Menu M0301090000 
            Caption         =   "Dar Visto Por Incumplimiento"
            Index           =   6
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
      Begin VB.Menu M0601000000 
         Caption         =   "Mantenimiento &Zonas"
         Index           =   3
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Mantenimiento Agencias"
         Index           =   4
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Mantenimiento de Ctas Contables"
         Index           =   5
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Copia de &Seguridad"
         Index           =   6
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Mantenimiento Grupo &Operaciones"
         Index           =   7
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Mantenimiento Operaciones Captaciones"
         Index           =   8
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Mantenimiento de Codigo Postal"
         Index           =   9
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Parámetros de Cheques"
         Index           =   10
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Mantenimiento CIIU"
         Index           =   11
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Mantenimiento de Feriados"
         Index           =   12
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Limites de Efectivo"
         Index           =   13
         Begin VB.Menu M0601010000 
            Caption         =   "Registro"
            Index           =   0
         End
         Begin VB.Menu M0601020000 
            Caption         =   "Mantenimiento"
            Index           =   1
         End
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Mantenimiento de Sesiones"
         Index           =   14
      End
   End
   Begin VB.Menu M0700000000 
      Caption         =   "&Personas"
      Index           =   0
      Begin VB.Menu M0701000000 
         Caption         =   "&Personas"
         Index           =   0
         Begin VB.Menu M0701010000 
            Caption         =   "&Registro"
            Index           =   0
         End
         Begin VB.Menu M0701010000 
            Caption         =   "&Mantenimiento"
            Index           =   1
         End
         Begin VB.Menu M0701010000 
            Caption         =   "&Consulta"
            Index           =   2
         End
         Begin VB.Menu M0701010000 
            Caption         =   "&Exoneradas del Lavado de Dinero"
            Index           =   3
         End
         Begin VB.Menu M0701010000 
            Caption         =   "&Roles de Persona"
            Index           =   4
         End
         Begin VB.Menu M0701010000 
            Caption         =   "&Comentarios a Persona"
            Index           =   5
         End
         Begin VB.Menu M0701010000 
            Caption         =   "&Grupos Economicos"
            Index           =   6
         End
         Begin VB.Menu M0701010000 
            Caption         =   "&Dudosas del Lavado de Dinero"
            Index           =   7
         End
         Begin VB.Menu M0701010000 
            Caption         =   "Registro Preventivo del &LAFT"
            Index           =   8
         End
         Begin VB.Menu M0701010000 
            Caption         =   "&Autorizacion de Clientes de Procedimiento Reforzado"
            Index           =   9
         End
         Begin VB.Menu M0701010000 
            Caption         =   "&Pre-Autorizacion de Clientes de Procedimiento Reforzado"
            Index           =   10
         End
         Begin VB.Menu M0701010000 
            Caption         =   "Administrar &Sesiones"
            Index           =   11
         End
         Begin VB.Menu M0701010000 
            Caption         =   "Montos Depositos de Ahorros"
            Index           =   12
         End
         Begin VB.Menu M0701010000 
            Caption         =   "Parametros Pago de Cuotas"
            Index           =   13
         End
         Begin VB.Menu M0701010000 
            Caption         =   "&Registro de Cliente Sensible"
            Index           =   14
         End
         Begin VB.Menu M0701010000 
            Caption         =   "Usuarios PREDA"
            Index           =   15
         End
         Begin VB.Menu M0701010000 
            Caption         =   "Promoción Actualiza tus Datos"
            Index           =   16
            Begin VB.Menu M0701010100 
               Caption         =   "Bitácora de Cambios"
               Index           =   0
            End
            Begin VB.Menu M0701010100 
               Caption         =   "Consulta de Bitácora"
               Index           =   1
            End
         End
         Begin VB.Menu M0701010000 
            Caption         =   "Clientes Honrados"
            Index           =   17
         End
         Begin VB.Menu M0701010000 
            Caption         =   "Búsqueda por Dirección"
            Index           =   18
         End
         Begin VB.Menu M0701010000 
            Caption         =   "Estados Financieros No Minoristas"
            Index           =   19
         End
         Begin VB.Menu M0701010000 
            Caption         =   "Visto de Continuidad del Proceso de Créditos de Personas del Reg. Prev. del LAFT"
            Index           =   20
         End
      End
      Begin VB.Menu M0701000000 
         Caption         =   "&Instituciones Financieras"
         Index           =   1
         Begin VB.Menu M0701020000 
            Caption         =   "Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu M0701020000 
            Caption         =   "Consulta"
            Index           =   1
         End
      End
      Begin VB.Menu M0701000000 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu M0701000000 
         Caption         =   "&Posicion de Cliente"
         Index           =   3
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
   Begin VB.Menu M100000000000 
      Caption         =   "A&uditoria"
      Index           =   0
      Begin VB.Menu M100100000000 
         Caption         =   "Unidad de Riesgos"
         Index           =   0
         Begin VB.Menu M100101000000 
            Caption         =   "Validación de la Calificación"
            Index           =   1
            Begin VB.Menu M100101010000 
               Caption         =   "Reporte de la Calificación de Cartera de Créditos"
               Index           =   2
            End
         End
         Begin VB.Menu M100102000000 
            Caption         =   "Ficha de Revisión"
            Index           =   3
            Begin VB.Menu M100102010000 
               Caption         =   "Registrar Ficha de Revisión"
               Index           =   4
            End
            Begin VB.Menu M100102020000 
               Caption         =   "Listar Ficha de Revisión"
               Index           =   5
            End
         End
         Begin VB.Menu M100103000000 
            Caption         =   "Pistas de Evaluación de Calificación"
            Index           =   6
         End
         Begin VB.Menu M100104000000 
            Caption         =   "Reporte de Cartera de Créditos - Auditoria"
            Index           =   7
         End
         Begin VB.Menu M100105000000 
            Caption         =   "Reporte de Garantías Inscritas"
            Index           =   8
         End
      End
      Begin VB.Menu M100200000000 
         Caption         =   "Ahorros"
         Index           =   9
         Begin VB.Menu M100201000000 
            Caption         =   "Cartas de Circularización - Saldos"
            Index           =   10
         End
         Begin VB.Menu M100202000000 
            Caption         =   "Reporte de las Operaciones de Ahorros"
            Index           =   11
         End
         Begin VB.Menu M100203000000 
            Caption         =   "Reporte de las Cuentas Aperturadas"
            Index           =   12
         End
         Begin VB.Menu M100204000000 
            Caption         =   "Movimientos de Cuentas de Ahorros"
            Index           =   13
         End
      End
      Begin VB.Menu M100300000000 
         Caption         =   "Créditos"
         Index           =   14
         Begin VB.Menu M100301000000 
            Caption         =   "Cartas de Circularización - Créditos"
            Index           =   15
         End
         Begin VB.Menu M100302000000 
            Caption         =   "Reporte de Operaciones Reprogramadas"
            Index           =   16
         End
         Begin VB.Menu M100304000000 
            Caption         =   "Reporte de Creditos Desembolsados - Cancelados"
            Index           =   17
         End
         Begin VB.Menu M100303000000 
            Caption         =   "Tarifario de Gastos"
            Index           =   18
            Begin VB.Menu M100303010000 
               Caption         =   "Consultar Tarifario de Gastos"
               Index           =   19
            End
            Begin VB.Menu M100303020000 
               Caption         =   "Reporte del Tarifario de Gastos"
               Index           =   20
            End
         End
      End
      Begin VB.Menu M100400000000 
         Caption         =   "Contabilidad"
         Index           =   21
         Begin VB.Menu M100401000000 
            Caption         =   "Balance del Mes Histórico"
            Index           =   22
         End
         Begin VB.Menu M100402000000 
            Caption         =   "Análisis de Cuentas"
            Index           =   23
         End
      End
      Begin VB.Menu M100500000000 
         Caption         =   "Logística"
         Index           =   24
         Begin VB.Menu M100501000000 
            Caption         =   "Movimientos de Proveedores"
            Index           =   25
            Begin VB.Menu M100501010000 
               Caption         =   "Orden Compra Soles"
               Index           =   26
            End
            Begin VB.Menu M100501020000 
               Caption         =   "Orden Compra Dolares"
               Index           =   27
            End
            Begin VB.Menu M100501030000 
               Caption         =   "Orden Servicio Soles"
               Index           =   28
            End
            Begin VB.Menu M100501040000 
               Caption         =   "Orden Servicio Dolares"
               Index           =   29
            End
         End
         Begin VB.Menu M100502000000 
            Caption         =   "Movimientos Proveedores X Fecha"
            Index           =   30
            Begin VB.Menu M100502010000 
               Caption         =   "Orden Compra Soles"
               Index           =   31
            End
            Begin VB.Menu M100502020000 
               Caption         =   "Orden Compra Dolares"
               Index           =   32
            End
            Begin VB.Menu M100502030000 
               Caption         =   "Orden Servicio Soles"
               Index           =   33
            End
            Begin VB.Menu M100502040000 
               Caption         =   "Orden Servicio Dolares"
               Index           =   34
            End
         End
      End
      Begin VB.Menu M100600000000 
         Caption         =   "Tesoreria"
         Index           =   35
         Begin VB.Menu M100601000000 
            Caption         =   "Reporte Pago a Proveedores"
            Index           =   36
         End
         Begin VB.Menu M100602000000 
            Caption         =   "Control de Saldos de Cuentas Corrientes"
            Index           =   37
         End
         Begin VB.Menu M100603000000 
            Caption         =   "Control de Saldos Adeudados"
            Index           =   38
         End
      End
      Begin VB.Menu M100700000000 
         Caption         =   "Sistemas"
         Index           =   39
         Begin VB.Menu M100701000000 
            Caption         =   "Relación de Accesos - Usuarios CMACMAYNAS S.A."
            Index           =   40
         End
      End
      Begin VB.Menu M100800000000 
         Caption         =   "Registro Auditoria"
         Index           =   41
         Begin VB.Menu M100801000000 
            Caption         =   "Registro Actividades Programadas"
            Index           =   42
         End
         Begin VB.Menu M100802000000 
            Caption         =   "Registro de Procedimientos"
            Index           =   43
         End
         Begin VB.Menu M100803000000 
            Caption         =   "Desarrollo de Procedimientos"
            Index           =   44
         End
         Begin VB.Menu M100804000000 
            Caption         =   "Verificar Procedimientos"
            Index           =   45
         End
      End
      Begin VB.Menu M100900000000 
         Caption         =   "Reportes"
         Index           =   46
         Begin VB.Menu M100901000000 
            Caption         =   "Seguimiento de Actividades"
            Index           =   47
         End
      End
   End
   Begin VB.Menu M110000000000 
      Caption         =   "Superv. Créditos"
      Enabled         =   0   'False
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu M110100000000 
         Caption         =   "Visitas a Clientes"
         Index           =   0
      End
      Begin VB.Menu M110200000000 
         Caption         =   "Control de Créditos"
         Index           =   1
         Begin VB.Menu M110201000000 
            Caption         =   "Pre-Desembolso [Control de Créditos]"
            Index           =   0
         End
         Begin VB.Menu M110201000000 
            Caption         =   "Post-Desembolso"
            Index           =   1
         End
         Begin VB.Menu M110201000000 
            Caption         =   "Configuración de CheckList"
            Index           =   2
         End
      End
      Begin VB.Menu M110300000000 
         Caption         =   "Reportes"
         Index           =   2
      End
      Begin VB.Menu M110400000000 
         Caption         =   "Mantenimiento Autorizaciones"
         Index           =   3
      End
      Begin VB.Menu M110500000000 
         Caption         =   "Mantenimiento Exoneraciones"
         Index           =   4
      End
      Begin VB.Menu M110600000000 
         Caption         =   "Hojas CF"
         Index           =   5
         Begin VB.Menu M110601000000 
            Caption         =   "Remesar"
            Index           =   0
         End
         Begin VB.Menu M110601000000 
            Caption         =   "Recepcionar"
            Index           =   1
         End
         Begin VB.Menu M110601000000 
            Caption         =   "Consultar"
            Index           =   2
         End
      End
      Begin VB.Menu M110700000000 
         Caption         =   "Créditos Vinculados"
         Index           =   6
         Begin VB.Menu M110701000000 
            Caption         =   "Asignación de Saldo[Créditos]"
            Index           =   0
         End
         Begin VB.Menu M110701000000 
            Caption         =   "Asignación de Saldo[Ventanilla]"
            Index           =   1
         End
         Begin VB.Menu M110701000000 
            Caption         =   "Estado de Saldo para Asignación"
            Index           =   2
         End
         Begin VB.Menu M110701000000 
            Caption         =   "Saldo Disponible por Colaborador"
            Index           =   3
         End
         Begin VB.Menu M110701000000 
            Caption         =   "Patrimonio Efectivo Ajustado"
            Index           =   4
         End
      End
      Begin VB.Menu M110700000000 
         Caption         =   "Mantenimiento CheckList"
         Index           =   7
         Begin VB.Menu M110801000000 
            Caption         =   "Autorizar"
            Index           =   0
         End
         Begin VB.Menu M110801000000 
            Caption         =   "Mantenimiento"
            Index           =   1
         End
         Begin VB.Menu M110801000000 
            Caption         =   "Consulta"
            Index           =   2
         End
      End
      Begin VB.Menu M110900000000 
         Caption         =   "Arqueos"
         Index           =   8
         Begin VB.Menu M110901000000 
            Caption         =   "Arqueo de Pagarés de Créditos"
            Index           =   0
         End
      End
   End
   Begin VB.Menu M120000000000 
      Caption         =   "Of. de Cumplimiento"
      Enabled         =   0   'False
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu M120100000000 
         Caption         =   "Parametros Agencia y Sector"
         Index           =   0
      End
      Begin VB.Menu M120200000000 
         Caption         =   "Operaciones Inusuales"
         Index           =   1
      End
      Begin VB.Menu M120300000000 
         Caption         =   "DJ Sujetos Obligados UIF"
         Index           =   2
      End
      Begin VB.Menu M120400000000 
         Caption         =   "Riesgo por Persona"
         Index           =   3
         Begin VB.Menu M120401000000 
            Caption         =   "Parámetros"
            Index           =   0
         End
      End
      Begin VB.Menu M120500000000 
         Caption         =   "Perfil por Actividad"
         Index           =   4
      End
   End
   Begin VB.Menu M130000000000 
      Caption         =   "Man&tenimiento"
      Index           =   0
      Begin VB.Menu M130100000000 
         Caption         =   "Mantenimiento de Giro"
         Index           =   0
      End
      Begin VB.Menu M130200000000 
         Caption         =   "Mantenimiento de Créditos Retirados"
         Index           =   1
      End
   End
   Begin VB.Menu M140000000000 
      Caption         =   "Unidad Seguros"
      Enabled         =   0   'False
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu M140100000000 
         Caption         =   "Seg. Tarjeta Débito"
         Index           =   0
         Begin VB.Menu M140101000000 
            Caption         =   "Configurar"
            Index           =   0
            Begin VB.Menu M140101010000 
               Caption         =   "Parámetros"
               Index           =   0
            End
            Begin VB.Menu M140101010000 
               Caption         =   "Nº Certificados por Agencia"
               Index           =   1
            End
            Begin VB.Menu M140101010000 
               Caption         =   "Documentos Act."
               Index           =   2
            End
         End
         Begin VB.Menu M140101000000 
            Caption         =   "Solicitud de Activacion"
            Index           =   1
         End
         Begin VB.Menu M140101000000 
            Caption         =   "Anulaciones"
            Index           =   2
         End
         Begin VB.Menu M140101000000 
            Caption         =   "Rechazo de Solicitud"
            Index           =   3
         End
         Begin VB.Menu M140101000000 
            Caption         =   "Aceptación de Solicitud"
            Index           =   4
         End
         Begin VB.Menu M140101000000 
            Caption         =   "Generación de Tramas"
            Index           =   5
         End
         Begin VB.Menu M140101000000 
            Caption         =   "Registrar Nota de Cargo"
            Index           =   6
         End
      End
      Begin VB.Menu M140100000000 
         Caption         =   "Seg ContraIncendio"
         Index           =   1
         Begin VB.Menu M140102000000 
            Caption         =   "Solicitud"
            Index           =   0
            Begin VB.Menu M140102010000 
               Caption         =   "Registro"
               Index           =   0
            End
            Begin VB.Menu M140102010000 
               Caption         =   "Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M140102010000 
               Caption         =   "Consulta"
               Index           =   2
            End
         End
         Begin VB.Menu M140102000000 
            Caption         =   "Respuesta"
            Index           =   1
            Begin VB.Menu M140102020000 
               Caption         =   "Rechazo de Solicitud"
               Index           =   0
            End
            Begin VB.Menu M140102020000 
               Caption         =   "Aceptación de Solicitud"
               Index           =   1
            End
            Begin VB.Menu M140102020000 
               Caption         =   "Reconsideración de rechazo"
               Index           =   2
            End
            Begin VB.Menu M140102020000 
               Caption         =   "Extorno"
               Index           =   3
            End
         End
         Begin VB.Menu M140102000000 
            Caption         =   "Consulta Historial"
            Index           =   2
         End
      End
      Begin VB.Menu M140100000000 
         Caption         =   "Configuración"
         Index           =   2
         Begin VB.Menu M140103000000 
            Caption         =   "Parámetros de Solicitud de Cobertura"
            Index           =   0
         End
      End
      Begin VB.Menu M140100000000 
         Caption         =   "Seguro Sepelio"
         Index           =   3
         Begin VB.Menu M140201000000 
            Caption         =   "Anulación"
            Index           =   0
         End
         Begin VB.Menu M140201000000 
            Caption         =   "Configuración"
            Index           =   1
            Begin VB.Menu M140201001000 
               Caption         =   "Parámetros"
               Index           =   0
            End
         End
         Begin VB.Menu M140201000000 
            Caption         =   "Solicitud de Activacíon"
            Index           =   2
         End
         Begin VB.Menu M140201000000 
            Caption         =   "Rechazo de Solicitud"
            Index           =   3
         End
         Begin VB.Menu M140201000000 
            Caption         =   "Aceptación de Solicitud"
            Index           =   4
         End
         Begin VB.Menu M140201000000 
            Caption         =   "Actualiza datos"
            Index           =   5
         End
      End
      Begin VB.Menu M140300000000 
         Caption         =   "Seg. MYPE"
         Index           =   4
         Begin VB.Menu M140300100000 
            Caption         =   "Generación de Trama"
            Index           =   0
         End
         Begin VB.Menu M140300100000 
            Caption         =   "Activación"
            Index           =   1
            Begin VB.Menu M140301100000 
               Caption         =   "Registro"
               Index           =   0
            End
            Begin VB.Menu M140301100000 
               Caption         =   "Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M140301100000 
               Caption         =   "Consulta"
               Index           =   2
            End
         End
         Begin VB.Menu M140300100000 
            Caption         =   "Respuesta"
            Index           =   2
            Begin VB.Menu M140300102000 
               Caption         =   "Rechazo Solicitud"
               Index           =   0
            End
            Begin VB.Menu M140300102000 
               Caption         =   "Aceptación Solicitud"
               Index           =   1
            End
            Begin VB.Menu M140300102000 
               Caption         =   "Reconsideración de Rechazo"
               Index           =   2
            End
            Begin VB.Menu M140300102000 
               Caption         =   "Extorno"
               Index           =   3
            End
         End
         Begin VB.Menu M140300100000 
            Caption         =   "Consulta Historial"
            Index           =   3
         End
      End
      Begin VB.Menu M140300000000 
         Caption         =   "Gestión de Siniestros"
         Index           =   5
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

Private Sub M0201000000_Click(Index As Integer)
    Select Case Index
        Case 4
            frmCapSimulacionPF.Show 1
            
        Case 9
           frmCapReportes.Show 1
           ' frmRepConsolidados.Show
           
        Case 10
            'frmAclAhorros.Show
            
        Case 11
            'frmAclColocaciones.Show
            
        Case 15
            frmExoneracionITF.Show
        
        Case 16
            frmCapNoCobroInactivas.Show 1
        Case 20
            frmCapBloqueoDesbloqueoParcial.inicia gCapAhorros
        Case 21
            frmCapPlazoFijoBloqueo.Show 1
        Case 22
            FrmCompraVentaAut.Inicio 'GIPO ERS0692016  07-01-2017
        Case 25
            'frmCapReasigGest.Show 1
        'MIOL RQ12257 ***************************************
        Case 31
            frmReqSunat.Show 1
        'END MIOL *******************************************
         'MIOL RQ12257 ***************************************
        'Case 32
         '   frmNivelesAprobacionCVAutorizar.InicioAutorizarNiveles
        'END MIOL *******************************************
    End Select
End Sub
'GITU 08-07-2016 ***************************
'Private Sub M0201010700_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmCapTarifarioGrupo.inicia gCapAhorros, False
'        Case 1
'            frmCapTarifarioGrupo.inicia gCapAhorros, True
'    End Select
'End Sub
'
'Private Sub M0201010800_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmCapTarifarioProgramacion.Show
'        Case 1
'            frmCapTarifarioProgramacion.Show
'    End Select
'End Sub
'
'Private Sub M0201010901_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmCapTarifarioComision.Show
'        Case 1
'            frmCapTarifarioComision.Show
'        Case 2
'    End Select
'End Sub
'END GITU **************************************

'JUEZ 20160415 **********************************
Private Sub M0201080000_Click(Index As Integer)
    Select Case Index
        Case 0
            frmCaptacCampanas.inicia
    End Select
End Sub
'END JUEZ ***************************************

'Private Sub M0201010000_Click(Index As Integer) Comentado por JUEZ 20130828
'    Select Case Index
'        Case 0
'             frmCapParametros.Inicia False
'        Case 1
'            frmCapParametros.Inicia True
'    End Select
'End Sub

'Private Sub M0201080000_Click(Index As Integer)
'  Select Case Index
'    Case 3
'         frmReporteSorteo.Show
'  End Select
'End Sub

Private Sub M0201080100_Click(Index As Integer)
  Select Case Index
    Case 0 'Mantenimiento
        'frmPreparaSorteo.Inicia ("00")
        frmCapTasaIntCamp.inicia gCapAhorros 'JUEZ 20160415
    Case 1 'Consulta
        'frmConsolidaSorteo.Inicia ("00")
        frmCapTasaIntCamp.inicia gCapAhorros, True 'JUEZ 20160415
  End Select
End Sub

Private Sub M0201080200_Click(Index As Integer)
Select Case Index
    Case 0 'Mantenimiento
        'frmPreparaSorteo.Inicia (gsCodAge)
        frmCapTasaIntCamp.inicia gCapPlazoFijo 'JUEZ 20160415
    Case 1 'Consulta
        'frmConsolidaSorteo.Inicia (gsCodAge)
        frmCapTasaIntCamp.inicia gCapPlazoFijo, True 'JUEZ 20160415
  End Select
End Sub
'JUEZ 20160415 ******************************************
Private Sub M0201080300_Click(Index As Integer)
    Select Case Index
        Case 0 'Mantenimiento
            frmCapTasaIntCamp.inicia gCapCTS
        Case 1 'Consulta
            frmCapTasaIntCamp.inicia gCapCTS, True
    End Select
End Sub
'END JUEZ ***********************************************

'JUEZ 20130828 **********************************
Private Sub M0201010100_Click(Index As Integer)
    Select Case Index
        Case 0
             frmCapParametros.inicia False
        Case 1
            frmCapParametros.inicia True
    End Select
End Sub

Private Sub M0201010200_Click(Index As Integer)
    Select Case Index
        Case 0
            'frmCapParametrosCom.Inicia 1 'Registro
            frmCapParametrosCom.inicia 1, "A" 'Registro 'JUEZ 20150930
        Case 1
            'frmCapParametrosCom.Inicia 2 'Consulta
            frmCapParametrosCom.inicia 2, "A" 'Consulta 'JUEZ 20150930
        Case 2
            'frmCapParametrosCom.Inicia 3 'Mantenimiento
            frmCapParametrosCom.inicia 3, "A" 'Mantenimiento 'JUEZ 20150930
    End Select
End Sub
'END JUEZ ***************************************

'RIRO20131212 ERS137
Private Sub M0201010300_Click(Index As Integer)
'    Dim oComision As New frmAdminComiTransBanc
'    Select Case Index
'        Case 0
'            oComision.nPermiso = 1
'            oComision.Show 1
'        Case 1
'            oComision.nPermiso = 2
'            oComision.Show 1
'    End Select
End Sub
'END RIRO

'RECO20140607 ERS008-2014***************************************************
Private Sub M0201010400_Click(Index As Integer)
    Select Case Index
        Case 0 'Mantenimiento
            frmGiroTarifarioMant.Inicio 1, "Giros - Tarifario - Mantenimiento"
        Case 1 'consulta
            frmGiroTarifarioMant.Inicio 2, "Giros - Tarifario - Consulta"
    End Select
End Sub
'RECO FIN*******************************************************************

Private Sub M0201010500_Click(Index As Integer)
Select Case Index
    Case 0 'Mantenimiento
            frmNivelesAprobacionCV.InicioRegistroNiveles
    Case 1 'Consulta
            frmNivelesAprobacionCVConsulta.InicioConsultaNiveles
End Select
End Sub

'JUEZ 20140908 ***********************************************/
Private Sub M0201010601_Click(Index As Integer)
    Select Case Index
        Case 0 'Mantenimiento
            frmCapParametros_NEW.Mantenimiento gCapAhorros
        Case 1 'Consulta
            frmCapParametros_NEW.Consulta gCapAhorros
    End Select
End Sub

Private Sub M0201010602_Click(Index As Integer)
    Select Case Index
        Case 0 'Mantenimiento
            frmCapParametros_NEW.Mantenimiento gCapPlazoFijo
        Case 1 'Consulta
            frmCapParametros_NEW.Consulta gCapPlazoFijo
    End Select
End Sub

Private Sub M0201010603_Click(Index As Integer)
    Select Case Index
        Case 0 'Mantenimiento
            frmCapParametros_NEW.Mantenimiento gCapCTS
        Case 1 'Consulta
            frmCapParametros_NEW.Consulta gCapCTS
    End Select
End Sub
'END JUEZ ****************************************************/

Private Sub M0201020201_Click(Index As Integer)
    Select Case Index
        Case 0 'Mantenimiento
            frmCapTasaInt.inicia gCapPlazoFijo, False
        Case 1 'Mantenimiento
            'frmCapTasaIntPF.Inicia gCapPlazoFijo, False
            frmCapTasaInt.inicia gCapPlazoFijo, False, True 'JUEZ 20140220
    End Select
End Sub

Private Sub M0201020202_Click(Index As Integer)
    Select Case Index
        Case 0 'Consulta
            frmCapTasaInt.inicia gCapPlazoFijo, True
        Case 1 'Consulta
            'frmCapTasaIntPF.Inicia gCapPlazoFijo, True
            frmCapTasaInt.inicia gCapPlazoFijo, True, True 'JUEZ 20140220
    End Select
End Sub


Private Sub M0201100000_Click(Index As Integer)

    Select Case Index
        Case 0 'Niveles
            FrmCapRegNivelAutRetCan.Show 1
        Case 1 'Niveles-Grupos
            FrmCapRegNivelAutRetCanDet.Show 1
        Case 2 'Aprobacion / rechazo
            'If Not VerificarRFIII Then Exit Sub ' *** RIRO SEGUN TI-ERS108-2013 ***
            FrmCapRegAproAutRetCan.Show 1
    End Select

End Sub

Private Sub M0201020100_Click(Index As Integer)
    Select Case Index
        Case 0 'mantenimiento
            frmCapTasaInt.inicia gCapAhorros, False
        Case 1 'Consulta
            frmCapTasaInt.inicia gCapAhorros, True
    End Select
End Sub

Private Sub M0201020200_Click(Index As Integer)
    Select Case Index
'        Case 0 'Mantenimiento
'            frmCapTasaInt.Inicia gCapPlazoFijo, False
'        Case 1 'Consulta
'            frmCapTasaInt.Inicia gCapPlazoFijo, True
        Case 2 'Cambio de Tasa Plazo Fijo
            frmCapCambioTasa.inicia gCapPlazoFijo
        'By Capi 07082008
        Case 3 'Cambio de Tasa Plazo Fijo
            frmCapCambioTasa.inicia gCapPlazoFijo, False
        '
    End Select
End Sub

Private Sub M0201020300_Click(Index As Integer)
    Select Case Index
        Case 0
            frmCapTasaInt.inicia gCapCTS, False
        Case 1
            frmCapTasaInt.inicia gCapCTS, True
        Case 2 'JUEZ 20140228
            frmCapCambioTasaCTSLote.Show 1
    End Select
End Sub

Private Sub M0201030000_Click(Index As Integer)
    'Mantenimiento
    Select Case Index
        Case 0 'Ahorro
            frmCapMantenimiento.inicia gCapAhorros
        Case 1 'Plazo fijo
            frmCapMantenimiento.inicia gCapPlazoFijo
        Case 2 'Cts
            frmCapMantenimiento.inicia gCapCTS
    End Select
End Sub

Private Sub M0201040000_Click(Index As Integer)
    'Bloqueos / Desbloqueos
    Select Case Index
        Case 0 'Ahorros
            'If Not VerificarRFIII Then Exit Sub ' *** RIRO SEGUN TI-ERS108-2013 ***
            frmCapBloqueoDesbloqueo.inicia gCapAhorros
        Case 1 'Plazo Fijo
            'If Not VerificarRFIII Then Exit Sub ' *** RIRO SEGUN TI-ERS108-2013 ***
            frmCapBloqueoDesbloqueo.inicia gCapPlazoFijo
        Case 2 'Cts
            'If Not VerificarRFIII Then Exit Sub ' *** RIRO SEGUN TI-ERS108-2013 ***
            frmCapBloqueoDesbloqueo.inicia gCapCTS
    End Select
End Sub

Private Sub M0201050000_Click(Index As Integer)
    Select Case Index
        Case 0 'Registro
            'frmTarjetaRegistra.Show 1
        Case 1 'Relacion
            'frmTarjetaBloqueo.Show 1
        Case 2 'Cambio de Clave
            'FrmTarjetaCambioClave.Show 1
    End Select
End Sub

Private Sub M0201060000_Click(Index As Integer)
    Select Case Index
        Case 0
            frmCapBeneficiario.inicia False
        Case 1
            frmCapBeneficiario.inicia True
    End Select
End Sub

Private Sub M0201070000_Click(Index As Integer)
    Select Case Index
        'Case 0 'generacion
            'frmCapOrdPagGenEmi.Show 1
        Case 1 'Certificacion
            frmCapOrdPagAnulCert.inicia gAhoOPCertificacion
        Case 2 'Anulacion
            frmCapOrdPagAnulCert.inicia gAhoOPAnulacion
        Case 3 'Consulta
            frmCapOrdPagConsulta.Show 1
        Case 4
            frmIngOP.Show 1
        Case 5
            'frmCartasOrdenes.Show 1
        'MIOL 20121006, SEGUN RQ12272 **********
        Case 6
            frmDesbloqueoClienteOrdPago.Show
        'END MIOL ******************************
    End Select
End Sub

Private Sub M0201070100_Click(Index As Integer)
Select Case Index
    Case 0
        frmCapOrdPagSolicitud.Show 1
    Case 1
        frmCapOrdPagEmiteImpr.Show 1
'        frmCapOrdPagProceso.Inicia gCapTalOrdPagEstSolicitado
'    Case 2
'        frmCapOrdPagProceso.Inicia gCapTalOrdPagEstEnviado
'    Case 3
'        frmCapOrdPagProceso.Inicia gCapTalOrdPagEstRecepcionado
End Select
End Sub



 

Private Sub M0201090000_Click(Index As Integer)
Select Case Index
    Case 0
        frmCapPersParam.inicia gCapAhorros
    Case 1
        frmCapPersParam.inicia gCapPlazoFijo
    Case 2
        frmCapPersParam.inicia gCapCTS
End Select
End Sub

Private Sub M0201110000_Click(Index As Integer)
Select Case Index
    Case 0
        frmCapTasaEspSeg.Show 1
    Case 1
        frmCapTasaEspAprRech.Show 1
    Case 2
        frmExtornoSolTasaEspecial.Show 1
    Case 4
        frmCapAdmNiveles.Show 1 ' Agregado Por RIRO el 20130418
        
End Select
End Sub

Private Sub M0201120000_Click(Index As Integer)
Select Case Index
    Case 0
        'frmCapCampanas.inicia
        'frmCapCampanas.Show 1
        'frmCapConvenioMant.Show 1
    Case 1
        'frmCapAsignaPremio.Inicia gCapPlazoFijo
        'frmCapAsignaPremio.Show 1
        'frmCapServConvCuentas.Show 1
    Case 2
        'frmCapServParametros.Show 1
    Case 4
        'frmCapServConvPlanPag.Show 1
End Select

End Sub

Private Sub M0201130000_Click(Index As Integer)
    Select Case Index
    Case 0 'Renovacion de Plazo
        FrmCapPlazoPanderito.Show 1
    End Select
End Sub
'MADM 20110318 - SERVICIOS
Private Sub M0201140000_Click(Index As Integer)
     Select Case Index
        Case 0
            'frmCredGeneraComisionPagoVarios.Show 1
        Case 1 'Gravar Garantias
            'frmCredGeneraLecturaFilePagoVarios.Show
        Case 2 'Solicitud CartaFianza
            'frmCredGeneraReportePagoVarios.Show 1
        
        'Agregado por RIRO el 20130418 ****
        Case 4
            frmCapCargaArchivo.Show 1
                
        Case 5
            frmCapGeneraArchivoRecaudo.Show 1
        ' Fin RIRO ****
        
    End Select
End Sub
'END MADM

' Agregado por RIRO el 20130418 *****
Private Sub M0201140100_Click(Index As Integer)

    Select Case Index
    
        Case 0
                frmCapRegistroConvenioAhorros.inicia 1
        
        Case 1
                frmCapRegistroConvenioAhorros.inicia 2
        
        Case 2
                frmCapRegistroConvenioAhorros.inicia 3
    
    End Select
    
End Sub
' Fin RIRO ****


'***Agregado por ELRO el 20130716, según RFC1306270002****
Private Sub M0201140200_Click(Index As Integer)
Select Case Index
    Case 0 'Registro de Convenio de Servicio de Pago
        frmCapServicioPago.iniciarRegistro
    Case 1 'Mantenimiento de Convenio de Servicio de Pago
        frmCapServicioPago.inicarMantenimiento
    Case 2 'Cargar Trama de Convenio de Servicio de Pago
        frmCapServicioPagoCargaArchivo.Show 1
    Case 3 'Baja Trama de Convenio de Servicio de Pago
        frmCapServicioPagoBajaArchivo.Show 1
End Select
End Sub
'***Fin Agregado por ELRO el 20130716, según RFC1306270002

'BRGO 20110425 - REGISTRO DE SUELDOS DE CLIENTES CTS
Private Sub M0201150000_Click(Index As Integer)
     Select Case Index
        Case 0 'Registro individual de sueldos CTS
            frmCapDatosSueldosCTS.Show 1
        Case 1 'Registro en lote de sueldos CTS
            frmCapCargaArchivoCTS.Show 1
        Case 2 'JUEZ 20140305
            frmCapCambioEstadoCTS.Show 1
    End Select
End Sub
'END BRGO

'Private Sub M0201160000_Click(Index As Integer)
'     Select Case Index
'        'JACA 20111025**********************************
'        Case 0 'Tarifario por Nro. de Operaciones Max.
'            frmCapTarifaOperaciones.Inicio 1
'        Case 1 'Consulta Tarifario por Nro. de Operaciones Max.
'            frmCapTarifaOperaciones.Inicio 2
'        Case 2 'Tarifario por Operaciones en Otras Agencias
'            frmCapParamComisionOtrasAge.Inicio 1
'        Case 3 'Consulta Tarifario por Operaciones en Otras Agencias
'            frmCapParamComisionOtrasAge.Inicio 2
'        'JACA END**************************************
'    End Select
'End Sub

Private Sub M0201170000_Click(Index As Integer)
    Select Case Index
    Case 0
        frmParametrosPIT.inicia False
    Case 1
        frmParametrosPIT.inicia True
    End Select
End Sub
'***Agregado por ELRO el 20120905, según OYP-RFC087-2012
Private Sub M0201180100_Click(Index As Integer)
Select Case Index
    Case 0: frmCapIndConCli.Show 1
    Case 1: frmCapIndConCliVoBo.Show 1
End Select
End Sub

Private Sub M0201180200_Click(Index As Integer)
    'frmCapIndComDep.Show 1
End Sub
'***Fin Agregado por ELRO el 20120905******************

Private Sub M0301010000_Click(Index As Integer)
    Select Case Index
    Case 0 'Solicitud CartaFianza
        frmCFSolicitud.Show 1
    Case 1 'Gravar Garantias
        'frmCredGarantCred.Inicio PorMenu, , 1
        frmGarantiaCobertura.Inicio InicioGravamenxMenu, CartaFianza
    Case 2 'Solicitud CartaFianza
        Call frmCFSugerencia.inicia
    Case 4 'Emitir CartaFianza
        frmCFEmision.Show 1
    Case 5 ' Honrar CartaFianza
        FrmCFHonrar.Show 1
    'Case 8 ' Niveles de Aprobacion
        'frmCFNivelApr.Show 1 YA NO SE USA
    
    Case 8 'Matenimiento de Tarifario
        'frmCFTarifario.Show 1
    Case 9  ' Relacionar con Credito
    
    Case 11  'Dar de Baja folio Hoja CF
    Case 12
    Case 13 'RECO20160310 ERS012-2016
        frmCFHojaAprob.Show 1
    End Select
End Sub

Private Sub M0301010100_Click(Index As Integer)
    Select Case Index
    Case 0 'Aprobacion
         frmCFAprobacion.Show 1
    Case 1 ' 'Rechazo
        FrmCFRechazar.Show 1
    Case 2
        FrmCFRetirarApr.Show 1
    Case 3
        FrmCFDevolucion.Show 1
    'By capi Set07
    Case 4
        FrmCFCancelacion.Show 1
   'MADM 20121201
    Case 5
        frmCFExtornoApro.Show 1
    'WIOR 20120828
    Case 6
        frmCFModEmision.Show 1
    'WIOR 20130311 ***************************
    Case 7: frmCFEditarModEm.Show 1
    'WIOR FIN ********************************
    End Select

End Sub

Private Sub M0301010200_Click(Index As Integer)
    Select Case Index
    Case 0 'Consultas
        frmCFHistorial.Show 1
        
    End Select
End Sub

Private Sub M0301010300_Click(Index As Integer)
    Select Case Index
    Case 0 'Reportes
        frmCFReporte.Show 1
    End Select

End Sub
'WIOR 20130312 ***************************************
Private Sub M0301010600_Click(Index As Integer)
 Select Case Index
    Case 0: frmCFAutRenovacion.Show 1
    Case 1: frmCFExtornoRenovacion.Inicio (1)
    Case 2: FrmCFRenovacion.Show 1
    Case 3: frmCFExtornoRenovacion.Inicio (2)
 End Select
End Sub
'WIOR FIN *******************************************
Private Sub M0301020000_Click(Index As Integer)
    Select Case Index
    Case 8 'Juez 20120905 '7 'Refinanciacion de Credito
        Call frmCredSolicitud.RefinanciaCredito(Registrar)
    Case 9 'Juez 20120905 '8 'Actualizacion con Metodos de Liquidacion
        frmCredMntMetLiquid.Show 1
    Case 10 'Juez 20120905 '9 'Perdonar Mora
        frmCredPerdonarMora.Show 1
    Case 12 'Juez 20120905 '11 'Reasignar Institucion
        frmCredReasigInst.Show 1
    Case 13 'Juez 20120905 '12 'Transferencia a Recuperaciones
        frmCredTransARecup.Show 1
    Case 18 'Juez 20120905 '17 'Registro de Dacion de Pago
        'frmCredRegisDacion.Show 1
    Case 19 'Juez 20120905 '18
        'frmCredCargoAuto.Show 1
    Case 20 'Juez 20120905 '19
        frmCredCodModular.Show 1
    Case 21 'Juez 20120905 '20
        frmCredAsigCComodin.Show 1
    Case 22 'Juez 20120905 '21
        'frmCredAdmPrepago.Show 1
    Case 23
        'frmCredValorizaCheque.Show 1
    ' CMACICA_CSTS - 05/11/2003 -------------------------------------------------
    Case 24
        'frmCredCalendarioDesemb.Show 1
    Case 25
        Call frmCredSolicitud.SustitucionCredito(Registrar)
    ' --------------------------------------------------------------------------
    'ALPA 20091007***********************************************
'    Case 36
         'Call frmCredSolicitud.AmpliacionCredito(Registrar)
    '************************************************************
    Case 27
        frmCredConvRegDev.Show 1
    Case 28
         'FrmVerRFA.Show vbModal
         
    Case 29
         'frmCredAutorizar.Show 1
    Case 30
        'frmCredCalendCOFIDE.Show 1
    Case 31
        'FrmCredCambioLC.Show vbModal
    Case 32
        'ARCV 14-02-2007
        Call frmCredSolicitud.AmpliacionCredito(Registrar)

         Case 35
            FrmCredGestionCobranza.Show 1
        'ALPA 20091007*****************************************
         Case 36
            frmGruposEconomicos.Show 1
        '******************************************************
        'JACA 20110628*****************************************
'        Case 39
'            frmCredRegistrarParametroBPPR.Show 1
        'JACA END******************************************
        
        Case 40 'JACA 20120109

         Case 47
            'FrmCredTraspCartera.Show 1
        'WIOR FIN *******************************************************
        Case 51
            frmCredSolicAutAmp.inicia
        Case 53
            frmCredAsistenteAgencia.Inicio 1, "Credito: Estado - Asistente Agencia", gTpoRegCtrlAsistAgencia  'ARLO20170925 ERS060-2016
        Case 54
            frmCredSolicAfp.Inicio '***CTI3 (ferimoro) ERS062-2018 29102018
            
    End Select
End Sub


Private Sub M0301020100_Click(Index As Integer)
        Select Case Index
            Case 10 'Registra Configuracion de Clientes Preferenciales
                frmCredConfClientesPreferenciales.Registrar
            Case 13 'JOEP ERS047
                frmCredRiesgoLimiteZonaGeog.inicia 1
            Case 14 'JOEP ERS047
                frmCredRiesgoLimiteporProducto.inicia 1
    End Select
End Sub

Private Sub M0301020101_Click(Index As Integer)
    Select Case Index
        Case 0 'Mantenimeinto de Parametros
            frmCredMantParametros.InicioActualizar
        Case 1 'Consulta de Parametros
            frmCredMantParametros.InicioCosultar
    End Select
End Sub

Private Sub M0301020102_Click(Index As Integer)
    Select Case Index
        Case 0 'Registro de Lineas de Credito
            'frmCredLineaCredito.Registrar
        Case 1 'Mantenimiento de Lineas de Credito
            'frmCredLineaCredito.actualizar
        Case 2 ' Consulta de lineas de Credito
            'frmCredLineaCredito.Consultar
    End Select
End Sub

Private Sub M0301020103_Click(Index As Integer)

    Select Case Index
        Case 0
            frmCredNewNivAprGrupoApr.InicioGrupoAprobacion
        Case 1
            frmCredNewNivAprGrupoApr.InicioRegistroNiveles
        Case 2
            frmCredNewNivAprParamApr.Show 1
        Case 3
            frmCredNewNivAprDelegacion.Inicio
    'End Select
    'END JUEZ ****************************************************
    'ALPA 20150215**************************
    Case 4
            frmParCargosTasa.Show 1
    End Select
    '***************************************
End Sub

Private Sub M0301020104_Click(Index As Integer)
    Select Case Index
        Case 0 'Mantenimeinto de Gastos
            frmCredMntGastos.Inicio InicioGastosActualizar
        Case 1
            frmCredMntGastos.Inicio InicioGastosConsultar
    End Select
End Sub

Private Sub M0301020105_Click(Index As Integer)
    Select Case Index
        Case 0
            FrmCampanas.Registro
        Case 1
            FrmCampanas.Mantenimiento
        Case 2
            FrmCampanas.Consultas
         Case 3
            frmCasaComercial.Show 'MADM 20100726
    End Select
    
End Sub

Private Sub M0301020106_Click(Index As Integer)
    Select Case Index
        Case 0 'JUEZ 20120905
            'frmCredEvalParamTipos.Show 1
        Case 1 'JUEZ 20120905
            'frmCredEvalParamIndicador.Show 1
        Case 2 'WIOR 20120905
            'frmCredEvalParamEspecializacion.Show 1
        Case 3 'PEAC 20160713
            frmCredFormEval.Show 1
        Case 4 'PEAC 20160713
            frmCredFormEvalConfigTpoProd.Show 1
    End Select
End Sub

' JUEZ 20121204 **********************************
Private Sub M0301020107_Click(Index As Integer)
    Select Case Index
        Case 0
            frmCredNewNivAprExoneracion.TiposExoneracion
        Case 1
            frmCredNewNivAprExoneracion.NivelesExoneracion
        Case 2 'RECO AGREGADO 20160909
            frmCredNewNivAutorizaConf.Show 1
        End Select
End Sub
' END JUEZ ***************************************

Private Sub M0301020108_Click(Index As Integer)
    Select Case Index
        Case 0 'Mantenimeinto
            'frmCredAdmPeriodo.InicioActualizar
        Case 1 'Consulta
            'frmCredAdmPeriodo.InicioConsultar
    End Select
End Sub
'WIOR 20131123 **************************************
Private Sub M0301020109_Click(Index As Integer)
    'ALPA 20150215****************
    Select Case Index
        Case 0
    '********************************
            'frmCredConfigCuotaBalon.Show 1
    'ALPA 20150215****************
        Case 1
            FrmCredLineaCreditoConfiguracion.Show 1
        Case 2
            frmCredConfigCatalogoProd.Inicio 'NAGL 20180711 Según ERS042-2018
    End Select
End Sub
'WIOR FIN *******************************************

'JUEZ 20150930 **************************************
Private Sub M0301020110_Click(Index As Integer)
    Select Case Index
        Case 0 'Registro
            frmCapParametrosCom.inicia 1, "C"
        Case 1 'Mantenimiento
            frmCapParametrosCom.inicia 3, "C"
        Case 2 'Consulta
            frmCapParametrosCom.inicia 2, "C"
    End Select
End Sub
'END JUEZ *******************************************

'JUEZ 20140530 *************************************
Private Sub M0301020111_Click(Index As Integer)
    Select Case Index
        Case 0 'Mantenimiento
            frmCredMantLimSecEcon.inicia 1
        Case 1 'Consulta
            frmCredMantLimSecEcon.inicia 2
        Case 2 'Configuracion Limites alertas tempranas 'LUCV20170302, Según ANEXO 001-2017
            frmCredAlertaTermpranaConfig.Show 1
    End Select
End Sub
'END JUEZ ******************************************
'FRHU 20160704 ERS002-2016
Private Sub M0301020112_Click(Index As Integer)
    Select Case Index
        Case 0 'Configurar Autorización
            'frmCredNewNivAutorizaConf.Show 1 'RECO20 COMENTADO 20160909
    End Select
End Sub
'FIN FRHU 20160704
Private Sub M0301020200_Click(Index As Integer)
    Select Case Index
        Case 0 'Registro de Solicitud
            frmCredSolicitud.Inicio Registrar
        Case 1 'Consulta de Solicitud
            frmCredSolicitud.Inicio Consulta
    End Select
End Sub

Private Sub M0301020300_Click(Index As Integer)
    Dim oCredRel As New UCredRelac_Cli  'COMDCredito.UCOMCredRela   'UCredRelacion

    Select Case Index
        Case 0 'Mantenimiento de Relaciones de Credito
            frmCredRelaCta.Inicio oCredRel, InicioMantenimiento, InicioRegistroForm
        '    Set oCredRel = Nothing
        Case 1 'Reasignacion de Cartera en Lote
            frmCredReasigCartera.Show 1
        Case 2 'Consulta de Relaciones de Credito
            frmCredRelaCta.Inicio oCredRel, InicioMantenimiento, InicioConsultaForm
        '    Set oCredRel = Nothing
        Case 3 'Confirmacion de reasignacion de cartera
            'frmCredRelaCta.Inicio oCredRel, InicioMantenimiento, InicioConsultaForm
            'frmCredConfirmaReasignaCartera.Show 1 'FRHU2014020 RQ14011
            frmCredConfirmaReasignaCartera.Inicio 'FRHU2014020 RQ14011
        Case 4 'FRHU 20140220 RQ14010 Asignacion de Agencia - Jefe de Negocios Territoriales
            frmAsigAgeJNTerritorial.Inicio
    End Select
    Set oCredRel = Nothing
End Sub

Private Sub M0301020400_Click(Index As Integer)
    Select Case Index
        Case 0 'Registro de Garantia

            frmGarantia.Registrar 'EJVG20150707
        Case 1 'Mantenimiento de Garantia
     
            frmGarantia.Editar 'EJVG20150707
        Case 2 'Consulta de Garantia

            frmGarantia.Consultar 'EJVG20150707
        Case 3 'Gravament
            'frmCredGarantCred.Inicio PorMenu
            frmGarantiaCobertura.Inicio InicioGravamenxMenu, Credito 'EJVG20150707
        Case 4 'Liberar Garantia
            'frmCredLiberaGarantia.Show 1 'EJVG20150707
        Case 5
            'FrmMantGarantias.Show vbModal
            frmGarantiaConf.Show 1
        Case 6
           'FrmCredRelGarantias.Show vbModal'EJVG20160309->Opción no se debe usar
        Case 7
           'FrmGraAmpliado.Show vbModal'EJVG20160309->Opción no se debe usar
        'MAVM 20100723 ***
        Case 8
           'frmJoyGarRegistro.Show vbModal'EJVG20160309->Opción no se debe usar
        Case 9
           'frmJoyGarRescate.Show vbModal'EJVG20160309->Opción no se debe usar
        '***
        'madm 20110420 ------------------
        Case 10
           frmCredGarantRealLegal.Show vbModal
      '----------------------------------
       'madm 20110525 ------------------
        Case 11
           frmCredGarantiaVerificaLegal.Show vbModal
      '----------------------------------
      'ALPA 20140318***********************************
        Case 12

      Case 13 'RECO20160213 ERS001-2016
            'frmGarantCartaVencimiento.Show 1'EJVG20160309->Coordinar luego del pase con ALPA para ver los cambios a realizar en esta opción.
        Case 14
            frmGarantiaCoberturaConfig.Show vbModal
        Case 15
            frmCredExoneraCobertura.Show vbModal
    End Select
End Sub

Private Sub M0301020500_Click(Index As Integer)
    'JUEZ 20121219 ******************
    Select Case Index
        Case 0 'Registro de Sugerencia
            If gnAgenciaCredEval = 0 Then
                '->***** LUCV20180601, Comentó y agregó según ERS022-2018
                'frmCredSugerencia.Sugerencia lSugerTipoActRegistro
                MsgBox "Agencia no configurada para el proceso de la sugerencia, por favor coordinar con TI", vbInformation, "Alerta"
                '<-***** Fin LUCV20180601
            Else
                'frmCredSugerencia_NEW.Sugerencia lSugerTipoActRegistro 'lSugerTipoActRegistro 'LUCV20180601, Comentó
                frmCredSugerencia_NEW.Sugerencia lSugerTipoActRegistroNew 'lSugerTipoActRegistro 'LUCV20180601, Agregó
            End If
        Case 1 'Consulta de Sugerencia
            If gnAgenciaCredEval = 0 Then
                '->***** LUCV20180601, Comentó y agregó según ERS022-2018
                'frmCredSugerencia.Sugerencia lSugerTipoActConsultar
                MsgBox "Agencia no configurada para el proceso de la sugerencia, por favor coordinar con TI", vbInformation, "Alerta"
                '<-***** Fin LUCV20180601
            Else
                'frmCredSugerencia_NEW.Sugerencia lSugerTipoActConsultar 'lSugerTipoActConsultar 'LUCV20180601, Comentó
                frmCredSugerencia_NEW.Sugerencia lSugerTipoActConsultarNew 'lSugerTipoActConsultar 'LUCV20180601, Agregó
            End If
        'WIOR 20160613 ***
         Case 2: frmCredDesbloqSobreEnd.Show 1
        'WIOR FIN ********
    End Select
    'END JUEZ ***********************
End Sub

Private Sub M0301020600_Click(Index As Integer)
    Select Case Index
        Case 0 'Aprobacion de Credito
            frmCredAprobacion.Show 1
        Case 1 'Rechazo de Credito
            frmCredRechazo.Rechazar
        Case 2 'Anulacion de Credito
            frmCredRechazo.Retirar
        Case 3
            FrmExtornoAprobacion.Show vbModal
        Case 4
            frmCredDesBloqCred.Show vbModal
        Case 5
            '20060329
            'En esta línea se rechaza la solictud de credito
            frmCredRechazo.Rechazar 3
        Case 6
            '20060329
            'en esta línea se rechaza la sugerencia de credito
            frmCredRechazo.Rechazar 4
        'JUEZ 20121204 ********************
        'MAVM 20110523 ***
        Case 7
            'frmCredPreAprobacion.Show vbModal
            frmCredNewNivAprPorNivel.Inicio
            
        Case 8
            'frmCredPreAprobacionListar.Show vbModal
            frmCredNewNivAprHist.Inicio
        '***
        'END JUEZ *************************
        'ALPA 20150215*********************
        Case 9
            frmCredCargosTasa.Show 1
        '**********************************
        Case 10 'EJVG20160606
            frmCredNewNivAprPorNivelExt.Show 1
        Case 11 'FRHU 20160702 ERS002-2016
            frmCredNewNivAutorizaVer.Inicio
        Case 12 'RECO20160713 ERS002-2016
            frmCredNewNivAprPorNivel.Inicio (nAprobacionAuto)
        'ALPA 20160714********************************************
        Case 13
            frmCredEndeuCuotaSistFinancTarjetas.Show 1
        '*********************************************************
    End Select
End Sub

Private Sub M0301020700_Click(Index As Integer)
    Select Case Index
        'JUEZ 20160216 *************************************
        Case 0
            frmCredReprogSolicitud.Inicio 0 'Solicitud
        Case 1
            frmCredReprogPropuesta.Inicio 0 'Propuesta
        Case 2
            'frmCredReprogSolicitud.Inicio 1 'Autorizacion 'Comentado JOEP20171214 ACTA220-2017
        Case 3
            frmCredReprogAprobacion.Inicio 0 'Aprobación
        Case 4
            frmCredReprogPropuesta.Inicio 1 'Rechazo
        Case 5 'Reprogramacion de Credito
            frmCredReprogCred.Show 1
        Case 6
            frmCredReprogAprobacion.Inicio 1 'VBAdmCred
        Case 7 'Reprogramacion en Lote
            frmCredReprogLote.Show 1
        Case 8
            'frmReestructuraRFA.Show 1
            frmCredReprogCredConvenio.Show 1 'WIOR 20140526
        'END IF ********************************************
        Case 9
            frmCredExtornoReprog.Show 1 'JOEP20170623
        Case 10
         frmCredReprogPropuesta.Inicio 2 'Add JOEP20210306 garantia covid
    End Select
End Sub

Private Sub M0301020800_Click(Index As Integer)
    Select Case Index
        Case 0 'Administracion de Gastos en Lote
            'frmCredAsigGastosLote.Show 1
        Case 1 ' mantenimiento de Penalidad
            frmCredExonerarPen.Show 1
        Case 2
            frmCredAdmiGastos.Show 1
        Case 3
            'frmColAsignacionGastoLote.Show 1
    End Select
End Sub

Private Sub M0301020900_Click(Index As Integer)
    Select Case Index
        Case 0 'Nota del Analista
            frmCredAsigNota.Show 1
        Case 1 'Meta del Analista
            'frmCredMetasAnalista.Show 1
            frmCredMetasAnalistas.Show 1 'JUEZ 20160407
    End Select
End Sub

Private Sub M0301021000_Click(Index As Integer)
    Dim MatCalend As Variant
    Dim Matriz(0) As String
    
    Select Case Index
        Case 0 'Calendario de Pagos
            frmCredCalendPagos.Simulacion DesembolsoTotal
        Case 1 'Desembolsos Parciales
            frmCredCalendPagos.Simulacion DesembolsoParcial
        Case 2 'Cuota Libre
'            MatCalend = frmCredCalendCuotaLibre.CalendarioLibre(True, gdFecSis, Matriz, 0#, 0, 0#)
        Case 3
            frmCredSimuladorPagos.Show 1
        Case 4
            'frmCredSimNroCuotas.Show 1
            frmCredSimuladorGarantiaPlazo.Show 1 'JUEZ 20140226
    End Select
End Sub

Private Sub M0301021100_Click(Index As Integer)
    If Index = 0 Then
        frmCredConsulta.Show 1
    Else
        'frmCredHistCalendario.Show 1
        Call frmCredHistCalendario.Inicio
    End If
End Sub

Private Sub M0301021200_Click(Index As Integer)
    Select Case Index
        Case 0
            frmCredDupDoc.Show 1
        Case 1
            frmCredReportes.inicia "Reportes de Créditos"
        Case 2
            frmCredVinculados.Ini True, "Créditos Vinculados - Titulares"
        Case 3
            frmCredVinculados.Ini False, "Créditos Vinculados - T y G Consolidado"
    End Select
End Sub



Private Sub M0301021300_Click(Index As Integer)
Select Case Index
        Case 0  'Registro para CrediPago
            'frmCredCrediPago.Show 1
            
        Case 1
            'frmCredCrediPagoArchivoCobranza.Show 1
        
        Case 2
            'frmCredCrediPagoArchivoResultado.Show 1
End Select
End Sub

Private Sub M0301021500_Click(Index As Integer)
Select Case Index
    Case 1
        Dim oGen As COMDConstSistema.DCOMGeneral   'DGeneral
        Dim lbCierreRealizado As Boolean
        
        Set oGen = New COMDConstSistema.DCOMGeneral   'DGeneral
        lbCierreRealizado = oGen.GetCierreDiaRealizado(gdFecSis)
        Set oGen = Nothing
        
        If lbCierreRealizado Then
            MsgBox "El cierre ya fue ralizado, no puede ingresar a esta opción.", vbInformation, "Aviso"
            Exit Sub
        End If
        frmCredCargoAuto.Show 1
    Case 2
        FrmMantCargoAutomatico.Show 1
End Select
End Sub

Private Sub M0301021600_Click(Index As Integer)
Select Case Index
    Case 1
        frmCredAdmPrepago.Show 1
    Case 2
        frmCredMntPrepago.Show 1
    Case 3 'JUEZ 20130925
        frmCredExonerarPreCanc.Show 1
End Select
End Sub

'EJVG20160310 ***
Private Sub M0301021701_Click(Index As Integer)

End Sub
'END EJVG *******
Private Sub M0301021800_Click(Index As Integer)
    Select Case Index
        Case 0 'Pagos en Banco de la Nación
            frmCredCorresponsaliaPagBcoNac.Show vbModal
        Case 1 'Desembolso en Banco de la Nación
            frmCredCorresponsaliaDesBcoNac.Show vbModal
        Case 2
            frmCredCorresponsaliaRecBcoNac.Show vbModal
        'MADM 20101020
        Case 3
            frmCredPagoConvenioBNreporte.Show vbModal
        'MADM ------
    End Select
End Sub

Private Sub M0301021900_Click(Index As Integer)
    Select Case Index
        Case 0
            frmCredComite.Show 1
        Case 1
            frmCredComiteReporte.Show 1
    End Select
End Sub

Private Sub M0301022100_Click(Index As Integer)
    'JACA 20110628*********************************
    Select Case Index
        Case 0
      
        Case 1
       
        Case 2

    End Select
    'JACA END****************************************
End Sub



Private Sub M0301022200_Click(Index As Integer)
    'BRGO 20111230*********************************
    Select Case Index
        Case 0
           
        Case 1
            frmCredPersonaCOFIDE.Inicio
        Case 2
            frmCredVehiculoInfoGas.Inicio
        Case 3
            frmINFOGASGeneracionXML.Show 1
        Case 4
            Call frmINFOGASLecturaArchivos.Inicio("05000", "Confirmación de habilitación de vehículo")
        Case 5
            Call frmINFOGASLecturaArchivos.Inicio("05001", "Recaudo de Vehículo")
    End Select
    'BRGO END****************************************
End Sub
'WIOR 20120613 **************************************************************
Private Sub M0301022300_Click(Index As Integer)
    Select Case Index
        Case 1
            'frmCredMicroseguro.Show 1
        Case 2
            'frmCredTramaMicroMult.Show 1
    End Select
End Sub
'WIOR FIN *******************************************************************
'MIOL 20120816, SEGUN RQ12077 ***********************************************
Private Sub M0301022400_Click(Index As Integer)
    Select Case Index
        Case 1
            'frmCredSugAprob.Inicio 1, "Credito: Estado - Asesoria Legal" 'RECO20161018 ERS060-2016
            frmCredSugAprob.Inicio 1, "Credito: Estado - Asesoria Legal", gTpoRegCtrlInformeLegal 'RECO20161018 ERS060-2016
        Case 2
            frmParametroRevisionExp.Show
        Case 3
            frmCredSugAprob.Inicio 2, "Credito: Estado - Supervición de Credito"
        Case 4 'RECO20161018 ERS060-2016
            'frmCredSugAprob.Inicio 1, "Credito: Estado - Asesoria Legal", gTpoRegCtrlMinutaLegal  'RECO20161018 ERS060-2016 'COMENTADO POR ARLO 20170919
            frmCredMinutaLegal.Inicio 1, "Crédito: Estado - Asesoría Legal - Minuta", gTpoRegCtrlMinutaLegal  'ARLO20170919 ERS060-2016
        Case 5 'RECO20161018 ERS060-2016
            frmCredSeguimiento.Show 1
    End Select
End Sub
'END MIOL *******************************************************************

'WIOR 20130411 ***********************************
Private Sub M0301022500_Click(Index As Integer)
Select Case Index
        Case 0: frmCredRiesgos.Show 1
        Case 1: frmCredRiesgosInformeMod.Show 1
        Case 2: frmConfirmaComunicacionCredRapiflash.inicia 'FRHU 20140331 RQ14178
        Case 4: frmCredRiesgosAutorizacionTipoCred.Inicio   'ARLO 2017030 ERS0652016
        Case 5: frmCredRiesgosAutorizacionMensual.Inicio    'ARLO 2017030 ERS0652016
        Case 6: frmCredRiesgosAutorizacionListado.Inicio    'ARLO 2017030 ERS0652016
        Case 7: frmCredRiesgoConfMonto.Show 1  'JOEP 20170610
        Case 8: frmCredRiesgoConfCatAge.Show 1 'JOEP 20170610
        Case 9: frmCredRiesgoConfgCodigosSobEnd.Show 1 'JOEP 20170810 ERS044
        Case 10: frmCredRiesgoLimiteZonaGeogAutoriza.Show 1 'JOEP 20170831 ERS047
        Case 11: frmCredRiesgoLimiteProductoAutoriza.Show 1 'JOEP 20170831 ERS047
End Select
End Sub
'WIOR FIN ****************************************

'JUEZ 20140530 ***********************************
Private Sub M0301022501_Click(Index As Integer)
Select Case Index
        Case 0: frmCredSolicLimSecEcon.Show 1
        Case 1: frmCredSolicLimSecEconExt.Show 1
End Select
End Sub
'END JUEZ ****************************************

'WIOR 20130727*************************************
Private Sub M0301022601_Click(Index As Integer)
Select Case Index
        Case 0: frmCredAgricoParam.Show 1
    End Select
End Sub
'WIOR FIN*************************************

'WIOR 20130924 ********************************
Private Sub M0301022700_Click(Index As Integer)
    Select Case Index
         'Case 0: frmCredRegVisitaAnalista.Show 1 'JUCS20210308
         'Case 1: frmCredRegVisitaJefe.Show 1 'JUCS20210308
    End Select
End Sub
'WIOR FIN *************************************
'RECO20140208 ERS002*******************************
Private Sub M0301022800_Click(Index As Integer)
      Select Case Index
        Case 0
            frmCredReasigInst.Show 1
        Case 1
            frmCredAsignarConvenio.Show 1
        Case 2
            frmCredRetiroConvenio.Show 1
        Case 3 'RECO20141018
            frmColHistoCambios.Show 1
    End Select
End Sub
'RECO FIN********************************************
'FRHU 20140324 ERS172-2013 RQ13875
Private Sub M0301022900_Click(Index As Integer)
    Select Case Index
        Case 0
            frmProyeccionSeguimientoPorAgencia.Show 1
        Case 1
            frmProyectadoVsEjecutadoPorAgencia.Show 1
    End Select
End Sub
'FIN FRHU 20140324 RQ13875

'WIOR 20160122 ***
Private Sub M0301023000_Click(Index As Integer)
    Select Case Index
        Case 0 'Administrador de Alertas Creditos MIVIVIENDA
            frmCredMiViviendaAlertasAdm.Show 1
    End Select
End Sub
'WIOR FIN ********

'JOEP ERS059-20161110
Private Sub M0301024000_Click(Index As Integer)
Select Case Index
        Case 0
            frmCredFichaSobreEndeudamiento.Inicio 1
        Case 1
            frmCredFichaSobreEndeudamiento.Inicio 2
        Case 2
            frmCredConsultarFichaSobreEnd.Inicio 3
    End Select
End Sub
'JOEP ERS059-20161110

'WIOR 20120905 **************************************************************
Private Sub M0301024500_Click(Index As Integer)
Select Case Index
        Case 0
            frmCredEvalSeleccion.Inicio 1
        Case 1
            frmCredEvalSeleccion.Inicio 2
        Case 2
            'frmCredEvalExtornoVerif.Show 1
            frmCredEvalSeleccion.Inicio 3 'EJVG20160712
        Case 3
            frmPersona.ConsultarFteIngreso "3010245" 'PTI120170530 segun ERS014-2017
    End Select
End Sub
'WIOR FIN *******************************************************************

Private Sub M0301030000_Click(Index As Integer)
    Select Case Index
        Case 1
            frmColPRescateJoyas.Show 1
        Case 3
        Case 5 ' Chafaloneo
        Case 11 'RECO20150219
            'frmColPCampAdjudicados.Show 1 'Comentado por TORE ERS054-2017
         Case 13 'APRI20180620 ERS063-2017
            frmRegHolograma.Show 1
        Case 14
            frmColPRetencion.RetencionCred 'JOEP2021 Campana Prendario
    End Select
End Sub

Private Sub M0301030100_Click(Index As Integer)
    Select Case Index
        Case 0 'Registro
            frmColPRegContratoDet.Show 1
        Case 1 'Mantenimiento
            frmColPMantPrestamoPig.Show 1
        Case 2 'Anulacion
            frmColPAnularPrestamoPig.Show 1
        Case 3 'Bloqueo
            'frmColPBloqueo.Show 1
        Case 4 'RECO20140208 ERS002**************************
            frmColPRegContratoAmpliacion.Show 1
    End Select
End Sub

'Private Sub M0301030200_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Entrega Joyas
'
'            frmColPRescateJoyas.Inicio gColPOpeDevJoyas, "Entrega de Joyas"
'        Case 1 'Entrega Joyas No Desembolsadas
'            frmColPRescateJoyas.Inicio gColPOpeDevJoyasNoDesemb, "Entrega Joyas No Desembolsadas"
'    End Select
'End Sub


Private Sub M0301030300_Click(Index As Integer)
'*** PEAC 20080714
'    Select Case Index
'        Case 0
'            frmColPRetasacion.Show 1
'        Case 1 'Preparacion de Remate
'            frmColPRematePrepara.Show 1
'        Case 2 'Remate
'            frmColPRemateProceso.Show 1
'    End Select
End Sub

Private Sub M0301030400_Click(Index As Integer)
    Select Case Index
        Case 0
            frmColPSubastaPrepara.Show 1
        Case 1
            frmColPSubastaProceso.Show 1
    End Select
End Sub

Private Sub M0301030500_Click(Index As Integer)
    
    Select Case Index
        Case 0
            'PEAC 20080605-descomentar para otra compilacion
            'frmColPRetasacion.Show 1
        Case 1
            frmColPVentaLotePrepara.Show vbModal
        Case 2
            frmColPVentaLoteProceso.Show vbModal
    End Select

End Sub

Private Sub M0301030600_Click(Index As Integer)
    Select Case Index
        Case 0
            frmColPMovs.Show 1
        Case 1
            frmColPContratosxCliente.Show 1
        Case 3
             frmColPRepo.Inicio 1
        Case 4
             frmColPRepo.Inicio 2
        Case 5
             frmColPRepo.Inicio 3
            
    End Select
End Sub

Private Sub M0301040000_Click(Index As Integer)
    Select Case Index
    Case 4
        'frmPigSeleccionVentaFundicion.Show 1
    Case 5
        'TORE 20210516, por memoria insuficiente.
        'frmPigConsulta.Inicio
    Case 6
        'frmPigFundicionJoya.Show 1
    Case 7
        'FrmPigRepValores.Show 1 ' X Memoria
    Case 8
        'frmPigEvaluacionMensualClientes.Show 1
    End Select
End Sub

Private Sub M0301040100_Click(Index As Integer)
    Select Case Index
    Case 0
        frmPigTarifario.Show 1
    Case 1
        'FrmPigClasificaCli.Show 1 'X Memoria
    End Select
End Sub

Private Sub M0301040200_Click(Index As Integer)
    Select Case Index
    Case 0
        'TORE comento 20210516, por falta de espacio
        'frmPigRegContrato.Show 1
    Case 1
        'frmPigMantContrato.Show 1
    Case 2
        'frmPigAnularContrato.Show 1 'X Memor
    Case 3
        'FrmPigBloqueo.Show 1 'X Memor
    End Select
End Sub

Private Sub M0301040300_Click(Index As Integer)
    Select Case Index
    Case 0
        'frmPigProyeccionGuia.Show 1
    Case 1
        'frmPigDespachoGuia.Show 1 'X Memom
    Case 2
        'frmPigRecepcionValija.Show 1
    End Select
End Sub

Private Sub M0301040400_Click(Index As Integer)

    Select Case Index
    Case 0
        'frmPigRegistroRemate.Show 1
    Case 1
        'frmPigProcesoRemate.Show 1
    Case 2
'        Dim oPigRemate As DPigContrato
'        Dim rs As Recordset
'
'        Set oPigRemate = New DPigContrato
'        Set rs = oPigRemate.dObtieneDatosRemate(oPigRemate.dObtieneMaxRemate() - 1)
'        If Not (rs.EOF And rs.BOF) Then
'            If CStr(rs!cUbicacion) = Right(gsCodAge, 2) Then
'                FrmPigVentaRemate.Show 1
'            Else
'                MsgBox "Usuario no se encuentra asignado en la Agencia de Remate", vbInformation, "Aviso"
'                Exit Sub
'            End If
'        End If
    End Select
End Sub

Private Sub M0301030700_Click(Index As Integer)
    Select Case Index
        Case 0 'prepara adjudicacion
            'PEAC 20080605-para desproteger en otra compilacion
            frmColPAdjudicaPrepara.Show 1
        Case 1 'adjudicacion de lotes
            frmColPAdjudicaLotes.Show 1
        Case 2
            frmColPReimpresionComprobanteAdj.Show 1
    End Select
End Sub

Private Sub M0301030800_Click(Index As Integer)
    Select Case Index
        Case 0
            frmPigMantenimientoPrecioTasacion.Show 1
        Case 1 'RECO20140721 ERS114*************
            frmColPTarifarioCartaNotarialMinka.Inicio 1, "Mantenimineto"
        'Case 2 'RECO20160215 ERS056-2015
            'frmColPTarifarioNotific.Show 1
    End Select
End Sub

'RECO20140805 ERS074-2014*************************
Private Sub M0301030900_Click(Index As Integer)
    Select Case Index
        Case 0
            frmColPPreparacionRetasacionVigDif.Show 1
        Case 1
            frmColPRetasacionConsulta.Show 1
        Case 2
            frmColPRetasacionVigenteDiferida.Show 1
    End Select
End Sub
'RECO FIN*****************************************

'JOEP20171211 ERS082-2017
Private Sub M0301031000_Click(Index As Integer)
    Select Case Index
        Case 0
            frmColPParmSegPrendario.Inicio
        Case 1
            frmColPRetencion.ConfigDatos 'JOEP20210927 campana Prendario
    End Select
End Sub
'JOEP20171211 ERS082-2017

'Private Sub M0301050000_Click(Index As Integer)
'    Select Case Index
        'Case 0 ' Ingreso a Recup de Otras Entidades
        '    frmColRecIngresoOtrasEnt.Show 1
        'Case 2 ' Gastos en Recuperaciones                          '***YIHU20190201 (MIGRAC. A SICMACM RECUP.)
        '    frmColRecGastosRecuperaciones.Show 1
        'Case 3 ' Metodo de Liquidacion
        '    frmColRecMetodoLiquid.Show 1                           '***YIHU20190201 (MIGRAC. A SICMACM RECUP.)
        'Case 5 ' Pago
        '    frmColRecPagoCredRecup.Inicio gColRecOpePagJudSDEfe, "PAGO CREDITO EN RECUPERACIONES", gsCodCMAC, gsNomCmac, True '***YIHU20190201 (MIGRAC. A SICMACM RECUP.)
        'Case 6 ' Cancelacion
        '   frmColRecCancelacion.Show 1  '***Juez 20120418
        '   frmColRecCancPagoJudicial.Show 1                        '***YIHU20190201 (MIGRAC. A SICMACM RECUP.)
        'Case 7 ' Castigo
        '    frmColRecCastigar.Show 1                               '***YIHU20190201 (MIGRAC. A SICMACM RECUP.)
        'Case 8
        
        'Case 10
            'frmGarLevant.Show 1'EJVG20160310-> Se comentó este Levantamiento xq por la misma opción de Garantías se realizará.
        'Case 11
            'frmGarantExtorno.Show 1'EJVG20160310-> Se comentó este Levantamiento xq por la misma opción de Garantías se realizará.
        'Case 13
        '    FrmColRecVistoRecup.Show 1                             '***YIHU20190201 (MIGRAC. A SICMACM RECUP.)
        'Case 14
            'frmCredTransfRecupeGarant.Show 1
        'Case 15
            'frmCredTransfGarantiaAdjudiSaneado.Show 1
        'Case 16
        '    frmColBienesAdjudLista.Show 1                          '***YIHU20190201 (MIGRAC. A SICMACM RECUP.)
        'Case 17
            'frmColEmbargadosListar.Show 1
        'Case 18
        '    FrmBloqueaRecupera.Show 1 'X Mem '***MADM 20111010     '***YIHU20190201 (MIGRAC. A SICMACM RECUP.)
        'Case 19
        '    FrmColRecRegVisitaCliente.Show 1 '*** PEAC 20120816    '***YIHU20190201 (MIGRAC. A SICMACM RECUP.)
            
'    End Select
'End Sub


'Private Sub M0301050100_Click(Index As Integer)                    '***YIHU20190201 (MIGRAC. A SICMACM RECUP.)
'    Select Case Index
'        Case 0
'            frmColRecExped.Show 1
'        Case 1
'            frmColRecActuacionesProc.inicia "N"
'    End Select
'End Sub

Private Sub M0301050200_Click(Index As Integer)
 Dim oCredRel As New UCredRelac_Cli 'COMDCredito.UCOMCredRela   'UCredRelacion
 Select Case Index
    Case 0
        frmCredRelaCta.Inicio oCredRel, InicioMantenimiento, InicioRegistroForm
        Set oCredRel = Nothing
    Case 1
        frmCredRelaCta.Inicio oCredRel, InicioMantenimiento, InicioConsultaForm
        Set oCredRel = Nothing
    'Case 2
    '    frmColRecComision.Show 1                                  '***YIHU20190201 (MIGRAC. A SICMACM RECUP.)
End Select
End Sub

'Private Sub M0301050300_Click(Index As Integer)                  '***YIHU20190201 (MIGRAC. A SICMACM RECUP.)
'Select Case Index
'    Case 0
'        frmColRecRConsulta.inicia "Consulta de Pagos de Créditos Judiciales"
'    Case 1
'        FrmColRecPagGestor.Show vbModal
'End Select
'End Sub

'Private Sub M0301050400_Click(Index As Integer)                   '***YIHU20190201 (MIGRAC. A SICMACM RECUP.)
'Select Case Index
'    Case 0
'        frmColRecReporte.inicia "Reportes de Recuperaciones"
'End Select
'End Sub

'Private Sub M0301050500_Click(Index As Integer)                '***YIHU20190201 (MIGRAC. A SICMACM RECUP.)
'Select Case Index
'    Case 0 ' Simulador
'        Call frmColRecNegCalculaCalendario.Inicio(1)
'    Case 1 'Registrar Negociacion
'        frmColRecNegRegistro.inicia (True)
'    Case 2 'Resolver Negociacion
'        frmColRecNegRegistro.inicia (False)
'End Select
'
'End Sub

''FRHU 20150428 ERS022-2015
'Private Sub M0301050600_Click(Index As Integer)                '***YIHU20190201 (MIGRAC. A SICMACM RECUP.)
'    Select Case Index
'        Case 0
'            frmColTransfCancelPago.Show 1
'    End Select
'End Sub
''FIN FRHU 20150428

''WIOR 20150602 ***
'Private Sub M0301052100_Click(Index As Integer)                '***YIHU20190201 (MIGRAC. A SICMACM RECUP.)
'Select Case Index
'    Case 0
'        frmRecupCampConfig.Show 1
'    Case 1
'        frmRecupCampAcoger.Show 1
'    Case 2
'        frmRecupCampAuto.Show 1
'End Select
'End Sub
''WIOR FIN ********

Private Sub M0301060000_Click(Index As Integer)
Select Case Index
    Case 1
        FrmColocCalCargaRCC.Show 1
    Case 3
        FrmColocEvalRep.Show , Me
    Case 4
        FrmColocCalConsultaCliente.Show 1
End Select
End Sub

Private Sub M0301060100_Click(Index As Integer)
Select Case Index
    Case 0
        frmColocCalEvalCli.Inicio True
    Case 1
        frmColocCalEvalCliAutomatico.Show 1
    Case 2 ' Este  formulario lo hizo Luis
        FrmColocCalGarantiasPreferidas.Show 1
    Case 3
        FrmColocCalReclasificados.Show 1
End Select
End Sub

Private Sub M0301060200_Click(Index As Integer)
Select Case Index
    Case 0
        'ARCV 04-06-2007
        'frmColocCalSist.Show 1
        frmColocCalSist_NEW.Show 1
    Case 1
        frmAudCierraCalificacion.Show 1
    Case 2
        frmColocCalTabla.Show 1
End Select

End Sub

Private Sub M0301060300_Click(Index As Integer)
    Select Case Index
        Case 1
            'FrmIntDevTpoCred.Show 1
        Case 2
            'frmDistrExposicion.Show 1
    End Select
End Sub

Private Sub M0301070000_Click(Index As Integer)
Select Case Index
    Case 0
        frmRCDParametro.Show 1
    Case 3
        'frmRCDReporte.Show , Me
        frmRCDReporte_NEW.Show , Me
        'frmColocCalEvalCli.Inicio True
    'ALPA 20120419***************
    Case 5
        frmRCDVcFecha.Show 1
    '****************************
    'ALPA 20120424***************
    Case 6
        frmRCDCambiarFormatoObs.Show 1
    '****************************
End Select

End Sub

Private Sub M0301070100_Click(Index As Integer)
Select Case Index
    Case 0 ' Datos Maestro RCD
         frmRCDActualizaRCDMaestroPersona.Show 1
    Case 1 ' Persona desde Maestro RCD
        frmRCDActPersDeMaestroPersona.Show 1
End Select
End Sub

Private Sub M0301070101_Click(Index As Integer)
Select Case Index
    Case 0 ' Carga de Errores
        frmRCDErrorCargaTXT.Show 1
    Case 1 ' Correccion de Errores
        frmRCDErrorCorreccion.Show 1
End Select
End Sub

Private Sub M0301070200_Click(Index As Integer)
Select Case Index
    Case 0 'Informe RCD
        'frmRCDGeneraDatosRCD.Show 1 'ARCV 06-06-2007
        frmRCDGeneraDatosRCD_NEW.Show 1
    Case 1 ' Informe IBM
        frmRCDGeneraDatosIBM.Show 1
    Case 2
        frmRCDVericaDatos.Show 1
End Select
End Sub

Private Sub M0301070300_Click(Index As Integer)
Select Case Index
    Case 0 ' RCDMaestro Persona
        frmRCDMantMaestroPersona.Show 1
End Select
End Sub

Private Sub M0301080000_Click(Index As Integer)
Select Case Index
    Case 0
        'frmCredSolicitud.Inicioleasing (Registrar) 'LUCV20171225, Comentó según MEMO 3143-2017
    Case 1
        'frmCredSugerencia.InicioCargaDatos Registrar, True'LUCV20171225, Comentó según MEMO 3143-2017
    Case 2
        'frmCredAprobacion.Aprobacion True 'LUCV20171225, Comentó según MEMO 3143-2017
    Case 3
        'frmCredDesembAbonoCta.DesembolsoCargoCuentaProveedorLeasing gCredDesembLeasing 'LUCV20171225, Comentó según MEMO 3143-2017
    Case 4
        'frmCredpagoCuotasLeasingDetalle.inicia gCredPagLeasingCI 'LUCV20171225, Comentó según MEMO 3143-2017
    Case 5
        'frmUsuarioLeasing.Show 1 'LUCV20171225, Comentó según MEMO 3143-2017
End Select
End Sub
'RECO20140208 ERS002***********************************************
Private Sub M0301090000_Click(Index As Integer)
    Select Case Index
        Case 0
            If Not gnAgenciaHojaRutaNew Then
                frmHojaRutaAnalista.Inicio 1, "- Registro"
            Else
                MsgBox "Ingrese a la opcion ''Colocaciones>Hoja Ruta>Resultado de Visita'' para registrar el resultado de tu visita (Agencia Configurada con la Nueva Hoja de ruta)."
            End If
        Case 1
            If Not gnAgenciaHojaRutaNew Then
                frmHojaRutaAnalista.Inicio 2, "- Mantenimiento" '& gsCodCargo
            Else
                MsgBox "Ingrese a la opcion ''Colocaciones>Hoja Ruta>Resultado de Visita'' para registrar el resultado de tu visita (Agencia Configurada con la Nueva Hoja de ruta)."
            End If
        Case 2
            If Not gnAgenciaHojaRutaNew Then
                frmHojaRutaAnalista.Inicio 3, "- Mantenimiento Coordinador" '& gsCodCargo
            Else
                MsgBox "Ingrese a la opcion ''Colocaciones>Hoja Ruta>Resultado de Visita'' para registrar el resultado de tu visita (Agencia Configurada con la Nueva Hoja de ruta)."
            End If
        Case 3
            If Not gnAgenciaHojaRutaNew Then
                frmHojaRutaAnalistaResultado.Inicio
            Else
                MsgBox "Ingrese a la opcion ''Colocaciones>Hoja Ruta>Resultado de Visita'' para registrar el resultado de tu visita (Agencia Configurada con la Nueva Hoja de ruta)."
            End If
        Case 4
            If Not gnAgenciaHojaRutaNew Then
                frmHojaRutaAnalistaConsulta.Show 1
            Else
                MsgBox "Ingrese a la opcion ''Colocaciones>Hoja Ruta>Resultado de Visita'' para registrar el resultado de tu visita (Agencia Configurada con la Nueva Hoja de ruta)."
            End If
        Case 5
            If gnAgenciaHojaRutaNew Then
                frmHojaRutaAnalistaGeneraResultado.Inicio
            Else
                MsgBox "Ingrese a las opciones anteriores para registrar el resultado de tu visita (Agencia No Configurada con la Nueva Hoja de ruta)."
            End If
        Case 6
            If gnAgenciaHojaRutaNew Then
                frmHojaRutaAnalistaDarVisto.Show 1
            Else
                MsgBox "Ingrese a las opciones anteriores para registrar el resultado de tu visita (Agencia No Configurada con la Nueva Hoja de ruta)."
            End If
    End Select
End Sub
'RECO FIN***********************************************************
'Private Sub M0401000000_Click(Index As Integer)
'If Index = 2 Or Index = 3 Or Index = 4 Or Index = 9 Then
'    Dim clsTC As COMDConstSistema.NCOMTipoCambio
'    Dim nTC As Double
'    Set clsTC = New COMDConstSistema.NCOMTipoCambio
'    nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
'    Set clsTC = Nothing
'    If nTC = 0 Then
'        MsgBox "NO se ha registrado el TIPO DE CAMBIO. Debe registrarse para iniciar operaciones.", vbInformation, "Aviso"
'        Exit Sub
'    End If
'End If
'
''Dim sfiltro() As String
''Dim lnFiltraTC As Integer
''Dim lnFiltraMP As Integer
'Dim oGen As COMDConstSistema.DCOMGeneral
'Dim lbCierreRealizado As Boolean
'Dim lbCierreCajaRealizado As Boolean
'Dim oCaj As COMNCajaGeneral.NCOMCajero
'
'Set oGen = New COMDConstSistema.DCOMGeneral
''lnFiltraTC = CInt(oGen.LeeConstSistema(102))
''lnFiltraMP = CInt(oGen.LeeConstSistema(103))
'lbCierreRealizado = oGen.GetCierreDiaRealizado(gdFecSis)
'Set oGen = Nothing
'
'If Not lbCierreRealizado Then
'    Set oCaj = New COMNCajaGeneral.NCOMCajero
'    lbCierreCajaRealizado = oCaj.YaRealizoCierreAgencia(gsCodAge, gdFecSis)
'    Set oCaj = Nothing
'End If
'    'RECO20151111 ERS061-2015******************
'        If lbCierreCajaRealizado Then
'            If Not VerificaGrupoPermisoPostCierre Then
'                lbCierreCajaRealizado = True
'            Else
'                lbCierreCajaRealizado = False
'            End If
'        End If
'    'RECO FIN *********************************
'    Select Case Index
'        Case 0
'            'frmMantTipoCambio.Show 1
'        Case 2
'            If lbCierreRealizado Then
'                MsgBox "El cierre ya fue realizado, no puede ingresar a esta opción.", vbInformation, "Aviso"
'                Exit Sub
'            End If
'            If lbCierreCajaRealizado Then
'                MsgBox "El cierre de caja de la agencia ya fue ralizado, no puede ingresar a esta opción.", vbExclamation, "Aviso"
'                Exit Sub
'            End If
''            ReDim sfiltro(5)
''            If lnFiltraMP = 1 Then
''                sfiltro(1) = "[1][01234][012]" 'Pigno Trujillo
''            ElseIf lnFiltraMP = 2 Then
''                sfiltro(1) = "[1][01345][012]"   'Pigno Lima
''            ElseIf lnFiltraMP = 0 Then
''                sfiltro(1) = "[1][012345][012]"   'Ambos
''            End If
''
''            sfiltro(2) = "[23][0-2][0123]"    'Captaciones
''
''            If lnFiltraTC = 0 Then
''                sfiltro(3) = "90002[0-3]"       'Compra Venta
''            ElseIf lnFiltraTC = 1 Then
''                sfiltro(3) = "90002[0-6]"
''            End If
''            sfiltro(4) = "9010[01][0123456789]"    'Control de Efectivo Boveda y Cajero
''            sfiltro(5) = "90003[0-5]"    'Operaciones con Cheques
''            frmCajeroOperaciones.Inicia sfiltro, "Cajero - Operaciones"
'            If gRsOpeF2.RecordCount = 0 Then
'                MsgBox "El usuario no tiene permisos para esta opción", vbInformation, "Mensaje"
'                Exit Sub
'            End If
'            'If Not VerificarRFIII Then Exit Sub ' *** RIRO SEGUN TI-ERS108-2013 ***
'            gRsOpeF2.MoveFirst
''            frmCajeroOperaciones.inicia "Cajero - Operaciones", gRsOpeF2
'
'        Case 3
'            If lbCierreRealizado Then
'                MsgBox "El cierre ya fue ralizado, no puede ingresar a esta opción.", vbInformation, "Aviso"
'                Exit Sub
'            End If
'            If lbCierreCajaRealizado Then
'                MsgBox "El cierre de caja de la agencia ya fue ralizado, no puede ingresar a esta opción.", vbExclamation, "Aviso"
'                Exit Sub
'            End If
''            ReDim sfiltro(4)
''            sfiltro(1) = "260[0-3]" 'Operaciones de Captaciones
''            sfiltro(2) = "126"      'Operaciones de Prendario
''            sfiltro(3) = "106"      'Operaciones de Colocaciones
''            sfiltro(4) = "136"      'Operaciones de judiciales
''            frmCajeroOpeCMAC.Inicia sfiltro, "Cajero - Operaciones CMACs Recepción"
'            If gRsOpeCMACRecep.RecordCount = 0 Then
'                MsgBox "El usuario no tiene permisos para esta opción", vbInformation, "Mensaje"
'                Exit Sub
'            End If
'            gRsOpeCMACRecep.MoveFirst
'            'frmCajeroOpeCMAC.inicia "Cajero - Operaciones CMACs Recepción", gRsOpeCMACRecep
'        Case 4
'            If lbCierreRealizado Then
'                MsgBox "El cierre ya fue ralizado, no puede ingresar a esta opción.", vbInformation, "Aviso"
'                Exit Sub
'            End If
'            If lbCierreCajaRealizado Then
'                MsgBox "El cierre de caja de la agencia ya fue ralizado, no puede ingresar a esta opción.", vbExclamation, "Aviso"
'                Exit Sub
'            End If
''            ReDim sfiltro(3)
''            sfiltro(1) = "2605"     'Operaciones de Captaciones
''            sfiltro(2) = "127"      'Operaciones de Prendario
''            sfiltro(3) = "107"      'Operaciones de Colocaciones
''            frmCajeroOpeCMAC.Inicia sfiltro, "Cajero - Operaciones CMACs Llamada"
'            If gRsOpeCMACLlam.RecordCount = 0 Then
'                MsgBox "El usuario no tiene permisos para esta opción", vbInformation, "Mensaje"
'                Exit Sub
'            End If
'            gRsOpeCMACLlam.MoveFirst
''            frmCajeroOpeCMAC.inicia "Cajero - Operaciones CMACs Llamada", gRsOpeCMACLlam
''            frmCajeroOpeCMAC.inicia "Cajero - Operaciones CMACs Recepción", gRsOpeCMACRecep
'        Case 5
'            If lbCierreRealizado Then
'                MsgBox "El cierre ya fue ralizado, no puede ingresar a esta opción.", vbInformation, "Aviso"
'                Exit Sub
'            End If
'            If lbCierreCajaRealizado Then
'                MsgBox "El cierre de caja de la agencia ya fue ralizado, no puede ingresar a esta opción.", vbExclamation, "Aviso"
'                Exit Sub
'            End If
'            If gRsOpeCMACLlam.RecordCount = 0 Then
'                MsgBox "El usuario no tiene permisos para esta opción", vbInformation, "Mensaje"
'                Exit Sub
'            End If
'            gRsOpeCMACLlam.MoveFirst
'           ' frmPITOperacionesInterCMAC.inicia "Cajero - Operaciones InterCMACs (Envío)" ', gRsOpeInterCMACs
'        Case 9
'            If lbCierreRealizado Then
'                MsgBox "El cierre ya fue realizado, no puede ingresar a esta opción.", vbInformation, "Aviso"
'                Exit Sub
'            End If
'            If lbCierreCajaRealizado Then
'                MsgBox "El cierre de caja de la agencia ya fue ralizado, no puede ingresar a esta opción.", vbExclamation, "Aviso"
'                Exit Sub
'            End If
''            ReDim sfiltro(9)
''            sfiltro(1) = "1[034][79][012]"         'Extornos de Colocaciones
''            If lnFiltraMP = 1 Then
''                sfiltro(2) = "129"          'Extornos de Prendario Trujillo
''            ElseIf lnFiltraMP = 2 Then      'Extornos de Prendario Lima
''                sfiltro(2) = "159"
''            ElseIf lnFiltraMP = 0 Then      'Extornos de Prendario Lima
''                sfiltro(2) = "1[25]9"
''            End If
''            sfiltro(3) = "2[3457]"      'Extornos de Captaciones
''            sfiltro(4) = "3[569]"       'Extornos de Otras Operaciones
''            If lnFiltraTC = 0 Then
''                sfiltro(5) = "90900[0-3]"
''            ElseIf lnFiltraTC = 1 Then
''               '  sfiltro(1) = "90900[0-6]"
''                sfiltro(5) = "90900[0-6]"
''            End If
''            sfiltro(6) = "90103[0-9]"   'Extornos de Operaciones de Boveda
''            sfiltro(7) = "90102[1-9]"   'Extornos de Operaciones de Cajero
''            sfiltro(8) = "90003[6-9]"   'Extornos de Operaciones con Cheque
''            sfiltro(9) = "90004[4-6]"   'Extornos de Compra Venta - Tipo de Cambio Especial
'
''            frmCajeroOperaciones.Inicia sfiltro, "Cajero - Extornos"
'            If gRsExtornos.RecordCount = 0 Then
'                MsgBox "El usuario no tiene permisos para esta opción", vbInformation, "Mensaje"
'                Exit Sub
'            End If
'            'If Not VerificarRFIII Then Exit Sub ' *** RIRO SEGUN TI-ERS108-2013 ***
'            gRsExtornos.MoveFirst
'            'frmCajeroOperaciones.inicia "Cajero - Extornos", gRsExtornos
''        Case 12
''
''              frmExoneracionITF.Show
''
''           ' AvisoOperacionesPendientes
''
''        Case 13
''             frmCapNoCobroInactivas.Show 1
''
'        '** Juez 20120723 *************
'         Case 16
'            'frmOpeReimprVoucher.Show 1
'        '** End Juez ******************
'
'        Case 18
'            'frmActivacionPerfilRFIII.Show 1 ' *** RIRO SEGUN TI-ERS108-2013 ***
'
'    End Select
'End Sub

Private Sub M0401010000_Click(Index As Integer)
Select Case Index
    Case 0
        'Case gOpeHabCajDevABoveMN, gOpeHabCajDevABoveME
        'frmCajeroHab.Show 1
        'Case gOpeHabCajTransfEfectCajerosMN, gOpeHabCajTransfEfectCajerosME
    Case 1
        'frmCajeroHab.Show 1
End Select
End Sub

Private Sub M0401000001_Click(Index As Integer)
Select Case Index
  Case 0
        'If M0401000001(0).Checked Then
         '   Timer1.Enabled = False
         '   M0401000001(0).Checked = False
        'ElseIf M0401000001(0).Checked = False Then
          '  Timer1.Enabled = True
          '  M0401000001(0).Checked = True
        'End If
  Case 1
       ' FrmCapAutOpeEstados.Show
  End Select
End Sub

Private Sub M0401030000_Click(Index As Integer)
    Select Case Index
        Case 0
'            Call frmCierreDiario.CierreDia
        Case 1
'            Call frmCierreDiario.CierreMes
        Case 2
        
        Case 3
        Case 4
            'FrmConsolidado.Show 1
    End Select
            
End Sub

Private Sub M0401040000_Click(Index As Integer)
    Select Case Index
        Case 0 'Captaciones
            'frmCapExtornos.Show 1
        
            
        Case 1 '    Extorno Credito
        
        Case 2 '    Pignoraticio
            'frmColPExtornoOpe.Show 1
        Case 3 '    recuperaciones
            
    End Select
End Sub

Private Sub M0401040100_Click(Index As Integer)
    Select Case Index
        Case 0
            'frmCredExtornos.Show 1
        Case 1
            'frmCredExtornos.Show 1
        Case 2
            'frmCredExtPagoLote.Show 1
    End Select
End Sub

Private Sub M0401050000_Click(Index As Integer)
    Select Case Index
    Case 0
        'frmAsientoDN.Inicio True
    Case 1
        'frmAsientoDN.Inicio False
    Case 2
        'frmCtaContMantenimiento.Show 1
    End Select
End Sub

Private Sub M0401060000_Click(Index As Integer)
Dim sCad As String
Dim oPrevio As previo.clsprevio

Select Case Index
    Case 0
        gsOpeDesc = "RESUMEN DE INGRESOS Y EGRESOS CONSOLIDADO"
        'frmCajeroIngEgre.inicia False, True
    Case 1
        Dim oRep As COMNCaptaGenerales.NCOMCaptaReportes
        
        Set oRep = New COMNCaptaGenerales.NCOMCaptaReportes
        'madm 20101012 - parametro agencia
        sCad = oRep.ReporteTrasTotSM("DETALLE DE OPERACIONES", False, gsCodUser, Format$(gdFecSis, "yyyymmdd"))
        Set oRep = Nothing
        
        Set oPrevio = New previo.clsprevio
        oPrevio.Show sCad, "DETALLE DE OPERACIONES", True
        Set oPrevio = Nothing
    Case 2
        Dim oHab As COMNCajaGeneral.NCajeroImp
        
        Usuario.Inicio gsCodUser
        Set oHab = New COMNCajaGeneral.NCajeroImp
        sCad = oHab.ReporteHabilitacionDevolucion(gsCodUser, Usuario.AreaCod, gsCodAge, gdFecSis, Usuario.UserNom, gsNomAge)
        
        Set oPrevio = New previo.clsprevio
        oPrevio.Show sCad, "DETALLE DE HABILITACIONES Y DEVOLUCIONES", True
        Set oPrevio = Nothing
    Case 3
        Dim oProt As COMNCaptaGenerales.NCOMCaptaReportes
        Set oProt = New COMNCaptaGenerales.NCOMCaptaReportes
        sCad = oProt.ProtocoloOperaciones("PROTOCOLO DE USUARIO SOLES", 0, 0, gsNomAge, gcEmpresa, gdFecSis, gMonedaNacional, gsCodUser, , Format(gdFecSis, gsFormatoFechaView), gsCodAge)
        sCad = sCad & oProt.ProtocoloOperaciones("PROTOCOLO DE USUARIO DOLARES", 0, 0, gsNomAge, gcEmpresa, gdFecSis, gMonedaExtranjera, gsCodUser, , Format(gdFecSis, gsFormatoFechaView), gsCodAge)
        
        Set oPrevio = New previo.clsprevio
        oPrevio.Show sCad, "PROTOCOLO DE USUARIO", True
        Set oPrevio = Nothing
    Case 4
        'frmOperacionesNum.Show 1
    Case 6
        'FrmITFGeneraArchivos.Show 1
    Case 8
        sCad = RepHavDevBoveda(gdFecSis, gsNomCmac, gsNomAge, gsCodAge)
        If sCad = "" Then
        MsgBox "No existe información"
        End If
    Case 9 '**DAOR 20080125, Reporte de Registros de Efectivo por usuario
        'frmOpeReportes.Show vbModal
        'MADM 20101019
    Case 10
        'frmCajeroIngDetalleGral.Show vbModal
    Case 11 'JUEZ 20130601
        'frmEnvioEstadoCtaRep.Show 1
    Case 12 'JUEZ 20131021
        'frmOpeRepActDatosCamp.Show 1
End Select
End Sub

Public Function RepHavDevBoveda(psFecSis As Date, psNomCmac As String, psNomAge As String, psCodAge As String) As String
  Dim cMovnro As String, rstemp As ADODB.Recordset
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
                xlHoja1.Cells(nFila, 6) = Format(CDate(Mid(rstemp!cMovnro, 5, 2) & "-" & Mid(rstemp!cMovnro, 7, 2) & "-" & Left(rstemp!cMovnro, 4)), "dd/MM/yyyy")
                xlHoja1.Cells(nFila, 7) = Mid(rstemp!cMovnro, 9, 2) & ":" & Mid(rstemp!cMovnro, 11, 2) & ":" & Mid(rstemp!cMovnro, 13, 2)
                                            
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
                xlHoja1.Cells(nFila, 6) = Format(CDate(Mid(rstemp!cMovnro, 5, 2) & "-" & Mid(rstemp!cMovnro, 7, 2) & "-" & Left(rstemp!cMovnro, 4)), "dd/MM/yyyy")
                xlHoja1.Cells(nFila, 7) = Mid(rstemp!cMovnro, 9, 2) & ":" & Mid(rstemp!cMovnro, 11, 2) & ":" & Mid(rstemp!cMovnro, 13, 2)
                                            
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



'Private Sub M0501020000_Click(Index As Integer)
'
'End Sub

 



'Private Sub M0401070000_Click(Index As Integer)
'Select Case Index
'    Case 1
'        'frmCajeroCierreAgencias.Show 1
'    Case 2
'        gsOpeCod = gOpeCajaExtCierreAgenica
'        gsOpeDesc = "EXTORNO CIERRE CAJA AGENCIA"
'        frmCajeroExtornos.inicia "CAJERO - EXTORNO CIERRE CAJA AGENCIA"
'End Select
'End Sub

'Private Sub M0401080000_Click(Index As Integer)
'Select Case Index
'    Case 0
'        frmCajeroBilletajeAutomatico.Inicio 1
'    Case 1
'        frmCajeroBilletajeAutomatico.Inicio 2
'End Select
'End Sub

'** Juez 20120807 ******************************
'Private Sub M0401090000_Click(Index As Integer)
'
'    'RIRO20140902 ***********
'    If Index = 0 Or Index = 1 Or Index = 3 Then
'        Dim clsTC As COMDConstSistema.NCOMTipoCambio
'        Dim nTC As Double
'        Set clsTC = New COMDConstSistema.NCOMTipoCambio
'        nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
'        Set clsTC = Nothing
'        If nTC = 0 Then
'            MsgBox "NO se ha registrado el TIPO DE CAMBIO. Debe registrarse para iniciar operaciones.", vbInformation, "Aviso"
'            Exit Sub
'        End If
'    End If
'    'END RIRO ***************
'
'    Select Case Index
'        Case 0
''            frmCajaArqueoVentBov.inicia 1 'Ventanilla
'        Case 1
'            'frmCajaArqueoVentBov.inicia 2 'Boveda
'        Case 2
'            'frmCajaArqueoVentBovExt.Show 1
'        'RIRO20140630 ERS072
'        Case 3
'            'frmCajaArqueoVentBov.inicia 3 'Entre Ventanillas
'        'END RIRO
'        Case 4 'PASI20151219
'            'frmArqueoTarjDebBoveda.inicia
'        Case 5
'            'frmArqueoTarjDebVent.inicia
'        Case 6
'            frmExtornoArqueoTarjDebBoveda.inicia 0
'        Case 7
'            frmExtornoArqueoTarjDebBoveda.inicia 1
'        'end PASI
'    End Select
'End Sub
'** End Juez ***********************************

'FRHU 20140505 ERS063-2014
'Private Sub M0401100000_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Aprobacion / Rechazo Operaciones
'            frmCapRegAproAutOtrasOperaciones.Show 1
'    End Select
'End Sub
'FIN FRHU 20140505

''RIRO 20150701 ERS162-2014
'Private Sub M0401110000_Click(Index As Integer)
'    Dim oUtilidades As New frmUtilidadesTrama
'    Select Case Index
'        Case 0
'            oUtilidades.inicia 1
'        Case 1
'            oUtilidades.inicia 2
'    End Select
'End Sub
'END RIRO

Private Sub M0601000000_Click(Index As Integer)
    Select Case Index
        Case 0 'Parametros
            
        Case 1 'Permisos responsabilidades
            '**DAOR 20071122 ******************
             frmMantPermisosResponsable.Show 1
            '**********************************
        Case 2 'permisos
            frmMantPermisos.Show 1
        'Case 3 'zonas
          '  frmMntZonas.Show 1
        'Case 4 'Agencias
         '   frmMntAgencias.Show 1
        Case 5 'Ctas Contables
            frmCtaContMantenimiento.Show 1
        Case 6 'backUp
            'frmBackUp.Show 1
        Case 7
            frmCajeroGrupoOpe.Show 1
        'Case 8
         '   frmCapMantOperacion.Show 1
        'Case 9
         '   frmMantCodigoPostal.Show 1
        Case 10
            frmDocRecParam.Show 1
        'Case 11
          '  frmMantCIIU.Show 1
        Case 12
            FrmMantFeriados.Show 1
        Case 14 'JUEZ 20160125
            frmMantSesiones.Show 1
    End Select
End Sub
'***Agregado por ELRO el 20120720, según OYP-RFC077-2012
Private Sub M0601010000_Click(Index As Integer)
    'frmLimEfeRegistro.Show 1
    frmLimiteEfectivoAdm.Inicio 1, "Registro" 'RECO20150206 ERS-022-2014
End Sub
'***Fin Agregado por ELRO el 20120720*******************

Private Sub M0601020000_Click(Index As Integer)
    'frmLimEfeMantenimiento.Show 1
    frmLimiteEfectivoAdm.Inicio 2, "Mantenimiento" 'RECO20150206 ERS-022-2014
End Sub

Private Sub M0701000000_Click(Index As Integer)
    If Index = 3 Then
        frmPosicionCli.Show 1
    End If
    'If Index = 6 Then  ' Reportes
    '    FrmPersReporte.Show 1
    'End If
End Sub

Private Sub M0701010000_Click(Index As Integer)
    'Persona
    Select Case Index
        Case 0 'Registro
            frmPersona.Registrar
        Case 1 'mantenimiento
            frmPersona.Mantenimeinto
        Case 2 'Consulta
            frmPersona.Consultar
        Case 3 'Exoneradas del Lavado de Dinero
            frmPersLavDinero.Show 1
        Case 4 'Rol de Persona
            FrmPersonaRolMantenimiento.Show 1
        Case 5
            frmPersComentario.Show 1
        Case 6
            frmPersGrupoE.Show 1
        'By Capi 30012008
        Case 7 'Dudosas del Lavado de Dinero
            frmPersLavDineroDudoso.Show 1
        
        '*** PEAC 20090715
        Case 8 'Registro Clientes lista negativa
            frmPersNegativas.Show 1
        'MADM 20100524 - Autorizacion Lista Negativa
        Case 9
            'frmPersNegativaAutorizacion.Show 1'WIOR 20121123 COMENTO
            frmPersNegativaAutorizacion.Inicio 'WIOR 20121123
    'WIOR 20121123 PRE AUTORIZACION ***********************************
       Case 10
            frmPersNegativaAutorizacion.Inicio True
    'WIOR FIN *********************************************************
        Case 11 'WIOR 20121123 cambio de 10 a 11
            frmPersAdministrarSesiones.Show 1
        Case 12 'WIOR 20121123 cambio de 11 a 12
            'frmCapAbonoPersParam.Show 1 'JACA  20110320 'FRHU 20171122 Comentado para compilar,su contenido del formulario esta comentado
        Case 13 'WIOR 20121123 cambio de 12 a 13
            frmCredPagoCuotasPersParam.Show 1 '** Juez 20120514
        Case 14 '** Modif Juez 20120514 Se cambio nro 12 >> 13'WIOR 20121123 cambio de 13 a 14
            frmPersClienteSensible.Show  'Modificado ELRO
        Case 15 'JUEZ 20130717
            frmPersPREDA.Inicio
        'WIOR 20140107 ******************
        Case 17
            frmCredHonrados.Show 1
        'WIOR FIN ***********************
        'FRHU 20140401 RQ14132 **********
        Case 18
            frmPersBusqueda.Show 1
        'FIN FRHU 20140401 **************
        Case 19 'FRHU 20150310 ERS013-2015
            frmPersEstadosFinancieros.Show 1
        Case 20 'MARG ERS046-2016
            frmPersRPLAFTVistoContinuidadCredito.Inicio
     End Select
End Sub

'JUEZ 20131016 *******************************************
Private Sub M0701010100_Click(Index As Integer)
    Select Case Index
        Case 0
            frmPersonaCampDatosBitacora.InicioActualizar
        Case 1
            frmPersonaCampDatosBitacora.InicioConsultar
    End Select
End Sub
'END JUEZ ************************************************

Private Sub M0701020000_Click(Index As Integer)
    'Instituciones Financieras
    Select Case Index
        Case 0
            frmMntInstFinanc.InicioActualizar
        Case 1
            frmMntInstFinanc.InicioConsulta
    End Select
End Sub

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

Private Sub M0901000000_Click(Index As Integer)
'frmRepMigra.Show 1
'frmInicioDia.Show 1
'frmCapRegeneraSaldos.Show 1
End Sub

'<****MAVM: Modulo de Auditoria 20/08/2008
Private Sub M100101010000_Click(Index As Integer)
'    FrmColocEvalRep.Show
'    FrmColocEvalRep.Inicializar_operacion
End Sub

Private Sub M100102010000_Click(Index As Integer)
    'frmRevisionRegistrar.Show 'FRHU 20171014 : Comentado por memoria llena
End Sub

Private Sub M100102020000_Click(Index As Integer)
    'frmListarRevision.Show 'FRHU 20171014 Comentado por Memoria Llena
End Sub

Private Sub M100103000000_Click(Index As Integer)
    'frmListarPista.Show 1
End Sub

Private Sub M100104000000_Click(Index As Integer)
    'frmComentarioCalificacion.Show 1
End Sub

Private Sub M100105000000_Click(Index As Integer)
    'frmCredReportes.inicia "Reportes de Créditos"
    'frmCredReportes.Inicializar_operacion ("108386")
End Sub

Private Sub M100201000000_Click(Index As Integer)
    frmGenerarCarta.Show 1
End Sub

Private Sub M100202000000_Click(Index As Integer)
'    frmCapReportes.Show
'    frmCapReportes.Inicializar_OperacionAhorros
End Sub

Private Sub M100203000000_Click(Index As Integer)
'    frmCapReportes.Show
'    frmCapReportes.Inicializar_CtaAperturadas
End Sub

'Private Sub M100204000000_Click(Index As Integer)
'    gRsOpeF2.MoveFirst
'    frmCajeroOperaciones.inicia "Cajero - Operaciones", gRsOpeF2
'    gRsOpeF2.MoveFirst
'    frmCajeroOperaciones.Inicializar_MovimientosCtasAhorros "Cajero - Operaciones", gRsOpeF2
'End Sub

Private Sub M100301000000_Click(Index As Integer)
    'frmGenerarCartaCredito.Show 1
End Sub

Private Sub M100302000000_Click(Index As Integer)
'    frmCredReportes.inicia "Reportes de Créditos"
'    frmCredReportes.Inicializar_OperacionesReprogramadas ("108325")
End Sub

Private Sub M100303010000_Click(Index As Integer)
    frmCredMntGastos.Inicio InicioGastosConsultar
End Sub

Private Sub M100303020000_Click(Index As Integer)
'    frmAuditReporteTarifario.Show 1
End Sub

Private Sub M100304000000_Click(Index As Integer)
    'frmAuditReporteCreditosDC.Show 1
End Sub

Private Sub M100401000000_Click(Index As Integer)
    'frmBalanceHisto.Inicio 1
End Sub

Private Sub M100402000000_Click(Index As Integer)
'    gbBitCentral = True
'    frmReportes.Inicio "76", 1
'    frmReportes.Inicializar_operacion
End Sub

Private Sub M100501010000_Click(Index As Integer)
    'gbBitCentral = True
    'frmLogOCAtencion.Inicio True, "501210", False, False, True
End Sub

Private Sub M100501020000_Click(Index As Integer)
    'gbBitCentral = True
    'frmLogOCAtencion.Inicio True, "502210", False, False, True
End Sub

Private Sub M100501030000_Click(Index As Integer)
    'gbBitCentral = True
    'frmLogOCAtencion.Inicio False, "501211", False, False, True
End Sub

Private Sub M100501040000_Click(Index As Integer)
    'gbBitCentral = True
    'frmLogOCAtencion.Inicio False, "502211", False, False, True
End Sub

Private Sub M100502010000_Click(Index As Integer)
    'gbBitCentral = True
    'frmLogOCAtencion.Inicio True, "501210", False, False, True, True
End Sub

Private Sub M100502020000_Click(Index As Integer)
    'gbBitCentral = True
    'frmLogOCAtencion.Inicio True, "502210", False, False, True, True
End Sub

Private Sub M100502030000_Click(Index As Integer)
    'gbBitCentral = True
    'frmLogOCAtencion.Inicio False, "501211", False, False, True, True
End Sub

Private Sub M100502040000_Click(Index As Integer)
    'gbBitCentral = True
    'frmLogOCAtencion.Inicio False, "502211", False, False, True, True
End Sub

Private Sub M100601000000_Click(Index As Integer)
'    gbBitCentral = True
'    frmReportes.Inicio "76", 2
'    frmReportes.Inicializar_Operacion_PagoProveedores
End Sub

Private Sub M100602000000_Click(Index As Integer)
'    gbBitCentral = True
'    frmReportes.Inicio "4[6-9]", 3
'    frmReportes.Inicializar_Operacion_ConsultaSaldos
End Sub

Private Sub M100603000000_Click(Index As Integer)
'    gbBitCentral = True
'    frmReportes.Inicio "4[6-9]", 4
'    frmReportes.Inicializar_Operacion_ConsultaSaldosAdeudos
End Sub

Private Sub M100701000000_Click(Index As Integer)
    'frmAuditListarAcceso.Show 1
End Sub

Private Sub M100801000000_Click(Index As Integer)
    'frmAudRegistroActividadProgramada.Show 1 'FRHU 20171014 Comentado por Memoria llena
End Sub

Private Sub M100802000000_Click(Index As Integer)
    'frmAudRegistroProcedimiento.Show 1 'FRHU 20171014 Comentado por Memoria llena
End Sub

Private Sub M100803000000_Click(Index As Integer)
    'frmAudDesarrolloProcedimiento.Show 1 'FRHU 20171014 Comentado por Memoria llena
End Sub

Private Sub M100804000000_Click(Index As Integer)
    'frmAudDesarrolloProcedimientoVerificar.Show 1 'FRHU 20171014 Comentado por Memoria llena
End Sub

Private Sub M100901000000_Click(Index As Integer)
'    frmAudReporteSeguimientoActividades.Show 1
End Sub

Private Sub M110100000000_Click(Index As Integer)
    'frmAdmCredRegVisitas.Show 1
End Sub

Private Sub M110200000000_Click(Index As Integer)
    'frmAdmCredRegControlCreditos.Show 1 'RECO20150421 ERS010-2015
End Sub

'RECO20150318 ERS010-2015************************
Private Sub M110201000000_Click(Index As Integer)
    Select Case Index
        Case 0
            'frmAdmControlCredDesemb.Show 1
        Case 1
            'frmAdmCredRegControl.Inicio "Post-Desembolso", "", 2
        'Case 2
           ' frmAdmConfigCheckList.Show 1
    End Select
End Sub
'RECO FIN***************************************


Private Sub M110300000000_Click(Index As Integer)
   'frmReportesAdmCred.Show 1
End Sub

Private Sub M110400000000_Click(Index As Integer)
    'frmAdmCredAutoMant.Show 1
End Sub

Private Sub M110500000000_Click(Index As Integer)
    'frmAdmCredExoMant.Show 1
End Sub
'WIOR 20120616 ************************************************
Private Sub M110601000000_Click(Index As Integer)
    Select Case Index
        'Hojas CF
        Case 0 'REMESAR
            'frmCFHojasRemesar.Show 1
        Case 1 'RECEPCIONAR
            'frmCFHojasRecepcion.Show 1
        Case 2 'CONSULTAR
            'frmCFHojasConsultar.Show 1
    End Select
End Sub
'WIOR FIN ***************************************************

'WIOR 20140128 ************************************************
Private Sub M110701000000_Click(Index As Integer)
  Select Case Index
        'Créditos Vinculados
        Case 0 'Asignacion Saldos[Créditos]
            'frmCredSaldosVincAsignar.Inicio (1)
        Case 1 'Asignacion Saldos[Ventanilla]
            'frmCredSaldosVincAsignar.Inicio (2)
        Case 2 'Estado Actual
            'frmCredSaldosVincEstado.Show 1
            'frmCredSaldosVincEstado.Inicio 'ORCR20140314
        Case 3 'Saldo Disponible
            'frmCredSaldosVincColaborador.Show 1
            'frmCredSaldosVincColaborador.Inicio 'ORCR20140314
        Case 4 'Patrimonio Efectivo Ajustado
            'frmPatrimonioEfectivo.Show 1
    End Select
End Sub
'WIOR FIN ***************************************************
'RECO20150326 ERS010-2015************************************
Private Sub M110801000000_Click(Index As Integer)
    Select Case Index
        Case 0
            'frmAdmCredAutorizacionChkList.Show 1
        Case 1
            'frmAdmCredChekListMant.Inicio "Mantenimiento", 1
        Case 2
            'frmAdmCredChekListMant.Inicio "Consulta", 2
    End Select
End Sub
'marg--
Private Sub M110901000000_Click(Index As Integer)
 Select Case Index
        Case 0
            'frmArqueoPagare.inicia
    End Select
End Sub
'end marg--

'RECO FIN ****************************************************
'JACA 20110514***********************************************
Private Sub M120100000000_Click(Index As Integer)
    'frmPersOpeAgeOcupacion.Show 1
End Sub
'JACA END****************************************************

Private Sub M120200000000_Click(Index As Integer)
    'JACA 20110530
    'frmOpeInusuales.Show 1 'FRHU 20171014 Comentado por Memoria llena
End Sub
'Comments PASI20171216 x Migracion a SICMACM Operaciones***********************
'Private Sub M120300000000_Click(Index As Integer)
'    frmDJSujetosObligados.Show 1
'End Sub
'end PASI*************************
'FRHU 20140917 ERS106-2014
'Comments PASI20171215 x Migracion a Operaciones
'Private Sub M120401000000_Click(Index As Integer)
'   frmOCNivelesRiesgoConfig.Show 1
'End Sub
'end PASI*************************
'Comments PASI20171216 x Migración a SICMACM Operaciones***********************
'Private Sub M120500000000_Click(Index As Integer)
'    frmOCNivelesRiesgoPorActividad.Show 1
'End Sub
'end PASI*************************
'FIN FRHU 20140917
Private Sub M130100000000_Click(Index As Integer)
    frmGiroMantenimiento.Show 1 '*** PEAC 20130222
End Sub

Private Sub M130200000000_Click(Index As Integer)
    frmMantCreditos.Show 1 '*** PEAC 20130222
End Sub
'FRHU 20140528 ERS068-2014
'COMENTADO POR ARLO20180604 POR MIGRACIÓN -  INICIO
'Private Sub M140101010000_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmSegTarjetaConfigParametros.Show 1
'        Case 1
'            frmSegTarjetaNCertificadosxAgencia.Show 1
'        Case 2
'            frmSegTarjetaConfigDoc.Show 1 'FRHU 20140610 ERS068-2014
'    End Select
'End Sub
'COMENTADO POR ARLO20180604 POR MIGRACIÓN - INICIO
'Private Sub M140101000000_Click(Index As Integer)
'    Select Case Index
'        Case 1
'            frmSegTarjetaSolicitud.Show 1 'FRHU 20140610 ERS068-2014
'        Case 2
'            frmSegTarjetaAfiliacionAnulacion.Show 1
'        Case 3
'            frmSegTarjetaRechazoSolicitud.Show 1 'FRHU 20140610 ERS068-2014
'        Case 4
'            frmSegTarjetaAceptacionSolicitud.Show 1 'FRHU 20140610 ERS068-2014
'        Case 5 'JUEZ 20140615
'            frmSegTarjetaGeneraTramas.Show 1
'        Case 6 'JUEZ 20150510
'            frmSegTarjetaAnulaDevPend.Show 1
'    End Select
'End Sub
'COMENTADO POR ARLO20180604 POR MIGRACIÓN - FIN

'FIN FRHU 20140528 ERS068-2014
'****>MAVM
'RECO20150326 ERS149-2014*************************
'COMENTADO POR ARLO20180604 POR MIGRACIÓN INICIO
'Private Sub M140102000000_Click(Index As Integer)
'    Select Case Index
'        Case 2
'            frmSegHistorialDocumentos.Show 1
'    End Select
'End Sub
'COMENTADO POR ARLO20180604 POR MIGRACIÓN FIN

'COMENTADO POR ARLO20180604 POR MIGRACIÓN INICIO
'Private Sub M140102010000_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmSegSolicitudCobertura.inicia 1, gContraIncendio 'RECO20160214 ERS073-2015
'        Case 1
'            frmSegSolicitudCobertura.inicia 2, gContraIncendio 'RECO20160214 ERS073-2015
'        Case 2
'            frmSegSolicitudCobertura.inicia 3, gContraIncendio 'RECO20160214 ERS073-2015
'    End Select
'End Sub

'Private Sub M140102020000_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmSegRechazoSolicitud.Show 1
'        Case 1
'            frmSegResolucionSolicCober.Show 1
'        Case 2
'            frmSegSolicReconsideracion.Show 1
'        Case 3
'            frmSegExtornoSolicCober.Show 1
'    End Select
'End Sub

'Private Sub M140103000000_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmSegParamSolicitud.Show 1
'    End Select
'End Sub
'COMENTADO POR ARLO20180604 POR MIGRACIÓN FIN

'RECO FIN*****************************************
'Private Sub M1000000000_Click(Index As Integer)
'    FrmCredTraspCartera.Show 1
'End Sub
'RECO20160213 ERS073-2015************************
'COMENTADO POR ARLO20180604 POR MIGRACIÓN INICIO
'Private Sub M140201000000_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmSegSepelioDesactivar.Show 1
'        Case 2
'            frmSegSolicitudCoberturaSepelio.inicia 1
'        Case 3
'            frmSegRechazoSolicitud.inicia gSepelio
'        Case 4
'            frmSegSepelioAceptacion.Show 1
'        Case 5 'RECO20160425
'            frmSegSepelioActDatos.Show 1
'    End Select
'End Sub

'Private Sub M140201001000_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmSegParamSepelio.Show 1
'    End Select
'End Sub

'WIOR 20160425 ***
'Private Sub M140300000000_Click(Index As Integer)
'    Select Case Index
'        Case 5: frmGestionSiniestro.Show 1
'    End Select
'End Sub
'WIOR FIN ********
'COMENTADO POR ARLO20180604 POR MIGRACIÓN FIN


'COMENTADO POR ARLO20180604 POR MIGRACIÓN INICIO
'Private Sub M140300100000_Click(Index As Integer)
'     Select Case Index
'        Case 0
'            'frmSegGeneracionTrama.Show
'        Case 3
'            frmSegHistorialDocumentos.inicia gMYPE
'    End Select
'End Sub
'Private Sub M140300102000_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmSegRechazoSolicitud.inicia gMYPE
'        Case 1
'            frmSegResolucionSolicCober.inicia gMYPE
'        Case 2
'            frmSegSolicReconsideracion.inicia gMYPE
'        Case 3
'            frmSegExtornoSolicCober.inicia gMYPE
'    End Select
'End Sub
'Private Sub M140301100000_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmSegSolicitudCobertura.inicia 1, gMYPE
'        Case 1
'            frmSegSolicitudCobertura.inicia 2, gMYPE
'        Case 2
'            frmSegSolicitudCobertura.inicia 3, gMYPE
'    End Select
'End Sub
'COMENTADO POR ARLO20180604 POR MIGRACIÓN FIN


'RECO FIN****************************************
Private Sub MDIForm_Click()
'  Form1.Show
End Sub

Private Sub MDIForm_Load()
 Me.Icon = LoadPicture(App.Path & gsRutaIcono)
Timer1.Enabled = False
CargaMensajes  'WIOR 20130826

'->***** LUCV20190323, Según RO-1000373 [Quita el borde de los dos controles]
Image1.BorderStyle = 0
pbxFondo.BorderStyle = 0
'<-***** Fin LUCV20190323

If gsCodCargo = "005002" Or gsCodCargo = "005003" Or gsCodCargo = "005004" Or gsCodCargo = "005005" Then

    'VAPI INTEGRAR BLOQUEO SICMACM PARA ANALISTAS QUE USAN EL APLICATIVO MOVIL 20160523
    Dim oDHojaRuta As DCOMhojaRuta
    Set oDHojaRuta = New DCOMhojaRuta
    
    If oDHojaRuta.participaHojaRuta(gsCodUser) Then
    
        Dim nRespuesta As Integer
        Dim bEsHoraLimite As Boolean
        bEsHoraLimite = oDHojaRuta.esHoraLimite()
        nRespuesta = oDHojaRuta.puedeGenerar(gsCodUser, bEsHoraLimite)
        Dim cMensaje As String
        Dim bEntrar As Boolean
        
        Select Case nRespuesta
            Case 1:
                oDHojaRuta.SolicitarVistoHojaRuta gsCodUser, gsCodAge, nRespuesta
                cMensaje = "Tiene resultados de días anteriores pendientes por enviar"
            Case 2:
                oDHojaRuta.SolicitarVistoHojaRuta gsCodUser, gsCodAge, nRespuesta
                cMensaje = "No ha generado su hoja de ruta en días anteriores"
            Case 3, 4:
                bEntrar = True
        End Select
        
        If (oDHojaRuta.tieneVistoPendiente(gsCodUser)) Then
            cMensaje = cMensaje & ",Solicita un visto bueno, regularice sus pendientes para poder entrar al sistema."
            bEntrar = False
        Else
            bEntrar = True
        End If
        
        If Not bEntrar Then
            MsgBox cMensaje, vbInformation, "AVISO"
           End
        End If
    'FIN VAPI
    End If

End If
'RECO20140124 ERS154********************************************************
'If gsCodCargo = "005002" Or gsCodCargo = "005003" Or gsCodCargo = "005004" Or gsCodCargo = "005005" Then
'
'    If Not 0 Then 'gnAgenciaHojaRutaNew Then
'    'Comentado por VAPI segun ERS0232015
'        Dim oCred As New COMDCredito.DCOMCreditos
'        Dim oDrNumEnDia As New ADODB.Recordset
'        Dim oDrNumDiaAnt As New ADODB.Recordset
'
'        Set oCred = New COMDCredito.DCOMCreditos
'        Set oDrNumEnDia = New ADODB.Recordset
'        Set oDrNumDiaAnt = New ADODB.Recordset
'
'        Set oDrNumEnDia = oCred.ObtieneNumeroRutaEnDia(Format(gdFecSis, "yyyy/MM/dd"), gsCodUser)
'        Set oDrNumDiaAnt = oCred.ObtieneNumeroRutaDiaAnt(Format(gdFecSis, "yyyy/MM/dd"), gsCodUser)
'        'Set oDrNumEnDia = oCred.ObtieneNumeroRutaEnDia(Format(gdFecSis, "yyyy/MM/dd"))
'        'Set oDrNumDiaAnt = oCred.ObtieneNumeroRutaDiaAnt(Format(gdFecSis, "yyyy/MM/dd"))
'
'        If Not (oDrNumDiaAnt.EOF And oDrNumDiaAnt.BOF) Then
'            If (oDrNumDiaAnt!nCantidad > 0) Then
'                MsgBox "Tiene pendiente registrar resultado de Hoja de Ruta Anterior", vbInformation, "AVISO"
'                frmHojaRutaAnalistaResultado.Inicio 1
'            End If
'        End If
'        If Not (oDrNumDiaAnt.EOF And oDrNumDiaAnt.BOF) Then
'            If (oDrNumEnDia!nCantidad > 0) Then
'            Else
'                MsgBox "Debe Registrar su Hoja de Ruta de forma obligatoria", vbInformation, "AVISO"
'                frmHojaRutaAnalista.Inicio 4, "Hoja de Ruta Analista"
'            End If
'        End If
'    'FIN COMENTADO POR VAPI
'    Else
'        'AGREGADO POR VAPI ERS 0232015
'        Screen.MousePointer = 0
'        'Dim oDhoja As New DCOMhojaRuta
'        Dim cPeriodo As String: cPeriodo = Format(gdFecSis, "YYYYMM")
'
''        If oDhoja.haConfiguradoAgencia(cPeriodo, gsCodAge) Then
''            Dim nPendiestesAtrasados As Integer: nPendiestesAtrasados = oDhoja.obtenerNumeroVisitasPendientes(gsCodUser, 0)
''            If nPendiestesAtrasados = 0 Then
''                Dim nPendiestes As Integer: nPendiestes = oDhoja.obtenerNumeroVisitasPendientes(gsCodUser, 1)
''                If nPendiestes > 0 Then
''                    If oDhoja.esHoraLimite Then
''                        frmHojaRutaAnalistaGeneraResultado.Show 1
''                    End If
''                Else
''                    If oDhoja.esHoraLimite Then
''                        frmHojaRutaAnalistaGenera.Inicio 1 'genera de mañana
''                    Else
''                        If oDhoja.obtenerNumeroVisitasRegistradasHoy(gsCodUser) <= 0 Then
''                            frmHojaRutaAnalistaGenera.Inicio 0 'genera de hoy
''                        End If
''                    End If
''                End If
''
''            Else
''                'ACA DEBE SOLICITAR EL VISTO DEL JEFE DE AGENCIA
''                If Not oDhoja.tieneVistoPendiente(gsCodUser) Then
''                    Dim cMovNro As String: cMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
''                    oDhoja.solicitarVisto gsCodUser, cMovNro
''                End If
''                If frmHojaRutaAnalistaVistoInc.Inicio Then
''                    frmHojaRutaAnalistaGeneraResultado.Show 1
''                Else
''                    End
''                End If
''            End If
''        Else
''            MsgBox "Aún no existe Configuración de Hoja de Ruta para la Agencia, comuníquelo al Jefe de Agencia."
''        End If
'    End If
'    'FIN VAPI
'End If
'RECO FINAL******************************************************************

'NO ES NECESARIO
' Dim SQLAUX As String, RsAUX As ADODB.Recordset, oconecta As DConecta
'
'
'Set oconecta = New DConecta
'oconecta.AbreConexion
'
'    SQLAUX = "   SELECT RC.CRHCARGOCOD FROM RRHH RH "
'    SQLAUX = SQLAUX & "  INNER JOIN RHCARGOS RC ON RC.CPERSCOD=RH.CPERSCOD "
'    SQLAUX = SQLAUX & " WHERE RH.CUSER='" & Vusuario & "'"
'
'    Set RsAUX = New ADODB.Recordset
'    RsAUX.CursorLocation = adUseClient
'    Set RsAUX = oconecta.CargaRecordSet(SQLAUX)
'
'
'    If RsAUX.Fields(0).value <> "006001" Then
'          Timer1.Enabled = False
'    End If
'    RsAUX.Close
'    Set RsAUX = Nothing
'
'oconecta.CierraConexion
'Set oconecta = Nothing

'FRHU20140319 RQ13874
Dim objCred As New COMDCredito.DCOMCredito
Dim objRS As ADODB.Recordset
Dim valor As Integer
Screen.MousePointer = 0
Set objRS = objCred.ValidarCargoProyeccionColocAge(gsCodCargo)
If objRS Is Nothing Then
    valor = 0
Else
    If Not objRS.EOF And Not objRS.BOF Then
        valor = objRS!valor
    End If
End If
If valor = 1 Then
    Call frmProyeccionPorAgencias.Inicio(gsCodAge, gdFecSis)
End If
Screen.MousePointer = 11
'FIN FRHU20140319 RQ13874

'->***** LUCV20190323, Según RO-1000373
    If VerificaGrupoMantenimientoUsuarios Then
        TlbMenu.Buttons.Item(6).Visible = True
        TlbMenu.Buttons.Item(6).Enabled = True
    Else
        TlbMenu.Buttons.Item(6).Visible = False
        TlbMenu.Buttons.Item(6).Enabled = False
    End If
'<-***** Fin LUCV20190323
End Sub


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
'Dim sSql As String, rs As ADODB.Recordset
'Dim oconecta As DConecta
'
'
'   On Error GoTo MensaError
'
'
'
'   sSql = "Select Sum(case when cautestado='A' then 1 else 0 end ) Aprobadas,Sum(case when cautestado='R' then 1 else 0 end ) Rechazadas,Sum(case when cautestado='P' then 1 else 0 end ) Pendientes "
'   sSql = sSql & " From capautorizacionope "
'   sSql = sSql & " where cautestado<>'E' and cuserori='" & Vusuario & "' and left(cultimaactualizacion,8)=convert(char(8),getdate(),112) "
'
'   Set oconecta = New DConecta
'   Set rs = New ADODB.Recordset
'     oconecta.AbreConexion
'     rs.CursorLocation = adUseClient
'     Set rs = oconecta.CargaRecordSet(sSql)
'     oconecta.CierraConexion
'     Set oconecta = Nothing
'     If rs.State = 1 Then
'          If (rs!Aprobadas > 0 Or rs!Rechazadas > 0 Or rs!Pendientes > 0) Then
'               Toolbar1.Visible = True
'               txtEstado1.Text = "Aprobadas: " & CStr(rs!Aprobadas)
'               txtEstado2.Text = "Rechazadas: " & CStr(rs!Rechazadas)
'               txtEstado3.Text = "Pendientes: " & CStr(rs!Pendientes)
'          Else
'               Toolbar1.Visible = False
'          End If
'          If rs.State = 1 Then rs.Close
'     End If
'        Set rs = Nothing
'
''  Dim i As Long
''       i = 0
''       For i = 0 To Timer1.Interval
''         If i = Timer1.Interval Then
''            '  Unload FrmCapAutOpeEstados
''            '  FrmCapAutOpeEstados.Show 1
''         End If
''         If i = (Timer1.Interval / 2) Then
''             '  Unload FrmCapAutOpeEstados
''         End If
''       Next i
'Exit Sub
'MensaError:
'     Call RaiseError(MyUnhandledError, "frmCapAutorizacionPrueba: CargaOperaciones  Method")
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
        '<-***** Fin LUCV20190323
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
    Call oSeguridad.InsertarPista(gIngresarSalirSistema, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gSalirSistema, "Salida del SICMACM Negocio" & " Versión: " & Format(App.Major, "#0") & "." & Format(App.Minor, "#0") & "." & Format(App.Revision, "#0") & "-" & frmLogin.gsFechaVersion)
     If oSeguridad.ValidaAccesoPistaRF(gsCodUser) Then
            Call oSeguridad.InsertarPistaSesion(gIngresarSalirSistema, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gSalirSistema, 1)
            Call oSeguridad.ActualizarPistaSesion(gsCodPersUser, GetMaquinaUsuario, 1) 'JUEZ 20160125
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
    frmSegMensajeMostrar.Inicio (Trim(rsSeg!cMensaje))
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
