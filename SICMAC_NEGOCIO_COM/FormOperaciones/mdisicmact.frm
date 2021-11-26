VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDISicmact 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema Operaciones"
   ClientHeight    =   10095
   ClientLeft      =   1500
   ClientTop       =   -1125
   ClientWidth     =   15615
   Icon            =   "mdisicmact.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pbxFondo 
      Align           =   1  'Align Top
      Height          =   9015
      Left            =   0
      Picture         =   "mdisicmact.frx":030A
      ScaleHeight     =   8955
      ScaleWidth      =   15555
      TabIndex        =   2
      Top             =   600
      Width           =   15615
      Begin VB.Image Image1 
         Height          =   13500
         Left            =   120
         Picture         =   "mdisicmact.frx":8602B
         Top             =   0
         Width           =   21600
      End
   End
   Begin SICMACT.Usuario Usuario 
      Left            =   600
      Top             =   720
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
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
            Picture         =   "mdisicmact.frx":10BD4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":10C066
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":10C380
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":10C69A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":10C9B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":10CCCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":10CE60
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":10D17A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":10E1CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":10F21E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":110270
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":1112C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisicmact.frx":112314
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
      Width           =   15615
      _ExtentX        =   27543
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
      Top             =   9765
      Width           =   15615
      _ExtentX        =   27543
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
            TextSave        =   "23/11/2021"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   "17:09"
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
   Begin VB.Menu M0400000000 
      Caption         =   "&Operaciones"
      Index           =   0
      Begin VB.Menu M0401000000 
         Caption         =   "Tipo de Ca&mbio"
         Index           =   0
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Operaciones"
         Index           =   2
         Shortcut        =   {F2}
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Operaciones CMACs &Recepción"
         Index           =   3
         Shortcut        =   {F3}
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Operaciones CMACs &Llamada"
         Index           =   4
         Shortcut        =   {F4}
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Operaciones InterCMACs"
         Index           =   5
         Shortcut        =   {F5}
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Cierres"
         Index           =   8
         Begin VB.Menu M0401030000 
            Caption         =   "Cierre de &Día"
            Index           =   0
         End
         Begin VB.Menu M0401030000 
            Caption         =   "Cierre de &Mes"
            Index           =   1
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "E&xtornos "
         Index           =   9
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Asiento Contable"
         Index           =   10
         Begin VB.Menu M0401050000 
            Caption         =   "Asiento Contable &Dia"
            Index           =   0
         End
         Begin VB.Menu M0401050000 
            Caption         =   "Asiento Contable &Anterior"
            Index           =   1
         End
         Begin VB.Menu M0401050000 
            Caption         =   "Mantenimiento &Asiento Contable"
            Index           =   2
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Reportes"
         Index           =   11
         Begin VB.Menu M0401060000 
            Caption         =   "Resumen de Ingresos y Egresos Consolidado"
            Index           =   0
         End
         Begin VB.Menu M0401060000 
            Caption         =   "Detalle de &Operaciones"
            Index           =   1
         End
         Begin VB.Menu M0401060000 
            Caption         =   "Detalle de &Habilitación/Devolución Cajero"
            Index           =   2
         End
         Begin VB.Menu M0401060000 
            Caption         =   "Protocolo por &Usuario"
            Enabled         =   0   'False
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu M0401060000 
            Caption         =   "Generar Archivo ITF"
            Index           =   6
         End
         Begin VB.Menu M0401060000 
            Caption         =   "Reporte Para Boveda de Hab/Dev"
            Index           =   8
         End
         Begin VB.Menu M0401060000 
            Caption         =   "Reporte de Registros de Efectivo por Usuario"
            Index           =   9
         End
         Begin VB.Menu M0401060000 
            Caption         =   "Detalle de Operaciones General"
            Index           =   10
         End
         Begin VB.Menu M0401060000 
            Caption         =   "Emisión de Estados de Cuentas"
            Index           =   11
         End
         Begin VB.Menu M0401060000 
            Caption         =   "Actualización de Datos por Campaña"
            Index           =   12
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Cierre de Caja de Agencia"
         Index           =   14
         Begin VB.Menu M0401070000 
            Caption         =   "&Cierre Caja Agencia"
            Index           =   1
         End
         Begin VB.Menu M0401070000 
            Caption         =   "&Extorno Cierre Caja Agencia"
            Index           =   2
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Generación Automática de Billetaje"
         Index           =   15
         Begin VB.Menu M0401080000 
            Caption         =   "&Generación"
            Index           =   0
         End
         Begin VB.Menu M0401080000 
            Caption         =   "&Extorno"
            Index           =   1
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Reimprimir Voucher"
         Index           =   16
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Arqueos"
         Index           =   17
         Begin VB.Menu M0401090000 
            Caption         =   "Registro Ventanilla"
            Index           =   0
         End
         Begin VB.Menu M0401090000 
            Caption         =   "Registro Bóveda"
            Index           =   1
         End
         Begin VB.Menu M0401090000 
            Caption         =   "Extornos"
            Index           =   2
         End
         Begin VB.Menu M0401090000 
            Caption         =   "Arqueos Entre Ventanilla"
            Index           =   3
         End
         Begin VB.Menu M0401090000 
            Caption         =   "Arqueo de Stock de Tarjetas de Débito - Bóveda"
            Index           =   4
         End
         Begin VB.Menu M0401090000 
            Caption         =   "Arqueo de Stock de Tarjetas de Débito - Ventanilla"
            Index           =   5
         End
         Begin VB.Menu M0401090000 
            Caption         =   "Extorno Arqueo de Stock de Tarjetas de Débito - Bóveda"
            Index           =   6
         End
         Begin VB.Menu M0401090000 
            Caption         =   "Extorno Arqueo de Stock de Tarjetas de Débito - Ventanilla"
            Index           =   7
         End
         Begin VB.Menu M0401090000 
            Caption         =   "Expedientes de Ahorros"
            Index           =   8
         End
         Begin VB.Menu M0401090000 
            Caption         =   "Extorno Arqueo Expedientes de Ahorros"
            Index           =   9
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Activación de Perfiles RFIII"
         Index           =   18
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Autorización"
         Index           =   19
         Begin VB.Menu M0401100000 
            Caption         =   "Aprobación/Rechazo Operaciones"
            Index           =   0
         End
         Begin VB.Menu M0401100000 
            Caption         =   "VB de atención sin Tarjeta"
            Index           =   1
         End
         Begin VB.Menu M0401100000 
            Caption         =   "Aprobar/Rechazar credito pignoraticio"
            Index           =   2
         End
         Begin VB.Menu M0401100000 
            Caption         =   "Pre Desembolso - Compra de Deuda."
            Index           =   3
         End
         Begin VB.Menu M0401100000 
            Caption         =   "Extorno Pre Desembolso - Compra de Deuda"
            Index           =   4
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Utilidades Ex Trabajadores "
         Index           =   20
         Begin VB.Menu M0401110000 
            Caption         =   "Carga de Trama"
            Index           =   0
         End
         Begin VB.Menu M0401110000 
            Caption         =   "Baja de Trama"
            Index           =   1
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Cámara de Compensación"
         Index           =   21
         Begin VB.Menu M0401120000 
            Caption         =   "Generación de Archivos de Envío"
            Index           =   0
         End
         Begin VB.Menu M0401120000 
            Caption         =   "Carga de Archivos"
            Index           =   1
         End
         Begin VB.Menu M0401120000 
            Caption         =   "Actualización de Oficinas "
            Index           =   2
         End
         Begin VB.Menu M0401120000 
            Caption         =   "Reporte de Saldos de Transferencia"
            Index           =   3
         End
         Begin VB.Menu M0401120000 
            Caption         =   "Extorno de Tramas de Presentados"
            Index           =   4
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Envío Estado de Cuenta"
         Index           =   22
         Begin VB.Menu M0401130000 
            Caption         =   "Afiliación Envío de Estado de Cuenta"
            Index           =   1
         End
         Begin VB.Menu M0401130000 
            Caption         =   "Desafiliación Envío de Estado de Cuenta"
            Index           =   2
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
         Caption         =   "Habilitación de Operaciones Especiales"
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
   Begin VB.Menu M1000000000 
      Caption         =   "Superv. Créditos"
      Index           =   0
      Begin VB.Menu M1001000000 
         Caption         =   "Visitas Clientes"
         Index           =   0
      End
      Begin VB.Menu M1002000000 
         Caption         =   "Control de Créditos"
         Index           =   0
         Begin VB.Menu M1002010000 
            Caption         =   "Pre-Desembolso [Control de Créditos]"
            Index           =   0
         End
         Begin VB.Menu M1002010000 
            Caption         =   "Post-Desembolso"
            Index           =   1
         End
         Begin VB.Menu M1002010000 
            Caption         =   "Configuración de CheckList"
            Index           =   2
         End
         Begin VB.Menu M1002010000 
            Caption         =   "Asignar Usuario PreDesembolso."
            Index           =   3
         End
      End
      Begin VB.Menu M1003000000 
         Caption         =   "Reportes"
         Index           =   0
      End
      Begin VB.Menu M1004000000 
         Caption         =   "Mantenimiento Autorizaciones"
         Index           =   0
      End
      Begin VB.Menu M1005000000 
         Caption         =   "Mantenimiento Exoneraciones"
         Index           =   0
      End
      Begin VB.Menu M1006000000 
         Caption         =   "Hojas CF"
         Index           =   0
         Begin VB.Menu M1006020000 
            Caption         =   "Remesar"
            Index           =   0
         End
         Begin VB.Menu M1006020000 
            Caption         =   "Recepcionar"
            Index           =   1
         End
         Begin VB.Menu M1006020000 
            Caption         =   "Consultar"
            Index           =   2
         End
      End
      Begin VB.Menu M1007000000 
         Caption         =   "Créditos Vinculados"
         Index           =   0
         Begin VB.Menu M1007030000 
            Caption         =   "Asignación de Saldo[Créditos]"
            Index           =   0
         End
         Begin VB.Menu M1007030000 
            Caption         =   "Asignación de Saldo[Ventanilla]"
            Index           =   1
         End
         Begin VB.Menu M1007030000 
            Caption         =   "Estado de Saldo para Asinación"
            Index           =   2
         End
         Begin VB.Menu M1007030000 
            Caption         =   "Saldo Disponible por Colaborador"
            Index           =   3
         End
         Begin VB.Menu M1007030000 
            Caption         =   "Patrimonio Efectivo Ajustado"
            Index           =   4
         End
      End
      Begin VB.Menu M1008000000 
         Caption         =   "Mantenimiento CheckList"
         Index           =   0
         Begin VB.Menu M1008040000 
            Caption         =   "Autorizar"
            Index           =   0
         End
         Begin VB.Menu M1008040000 
            Caption         =   "Mantenimiento"
            Index           =   1
         End
         Begin VB.Menu M1008040000 
            Caption         =   "Consulta"
            Index           =   2
         End
      End
      Begin VB.Menu M1009000000 
         Caption         =   "Arqueos"
         Index           =   0
         Begin VB.Menu M1009050000 
            Caption         =   "Arqueo de Pagarés de Créditos"
            Index           =   0
         End
      End
   End
   Begin VB.Menu M1100000000 
      Caption         =   "Of. de Cumplimiento"
      Index           =   0
      Begin VB.Menu M1101000000 
         Caption         =   "Parametros Agencia y Sector"
         Index           =   0
      End
      Begin VB.Menu M1102000000 
         Caption         =   "Operaciones Inusuales"
         Index           =   0
      End
      Begin VB.Menu M1103000000 
         Caption         =   "DJ Sujetos Obligados UIF"
         Index           =   0
      End
      Begin VB.Menu M1104000000 
         Caption         =   "Riesgo por Persona"
         Index           =   0
         Begin VB.Menu M1104010000 
            Caption         =   "Parámetros"
            Index           =   0
         End
      End
      Begin VB.Menu M1105000000 
         Caption         =   "Perfil por Actividad"
         Index           =   0
      End
   End
   Begin VB.Menu M1200000000 
      Caption         =   "Unidad Seguros"
      Index           =   0
      Begin VB.Menu M1201000000 
         Caption         =   "Seg. Tarjeta Débito"
         Index           =   0
         Begin VB.Menu M1201010000 
            Caption         =   "Configurar"
            Enabled         =   0   'False
            Index           =   0
            Begin VB.Menu M1201010100 
               Caption         =   "Parámetros"
               Enabled         =   0   'False
               Index           =   0
            End
            Begin VB.Menu M1201010100 
               Caption         =   "N° Certificados por Agencia"
               Enabled         =   0   'False
               Index           =   1
            End
            Begin VB.Menu M1201010100 
               Caption         =   "Documentos Act."
               Enabled         =   0   'False
               Index           =   2
            End
         End
         Begin VB.Menu M1201010000 
            Caption         =   "Solicitud de Activación"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu M1201010000 
            Caption         =   "Anulaciones"
            Index           =   2
         End
         Begin VB.Menu M1201010000 
            Caption         =   "Rechazo de Solicitud"
            Enabled         =   0   'False
            Index           =   3
         End
         Begin VB.Menu M1201010000 
            Caption         =   "Aceptación de Solicitud"
            Enabled         =   0   'False
            Index           =   4
         End
         Begin VB.Menu M1201010000 
            Caption         =   "Generación de Tramas"
            Enabled         =   0   'False
            Index           =   5
         End
         Begin VB.Menu M1201010000 
            Caption         =   "Registrar Nota de Cargo"
            Index           =   6
         End
      End
      Begin VB.Menu M1202000000 
         Caption         =   "Seg. Contra Incendio"
         Enabled         =   0   'False
         Index           =   0
         Begin VB.Menu M1202010000 
            Caption         =   "Solicitud"
            Enabled         =   0   'False
            Index           =   0
            Begin VB.Menu M1202010100 
               Caption         =   "Registro"
               Enabled         =   0   'False
               Index           =   0
            End
            Begin VB.Menu M1202010100 
               Caption         =   "Mantenimiento"
               Enabled         =   0   'False
               Index           =   1
            End
            Begin VB.Menu M1202010100 
               Caption         =   "Consulta"
               Enabled         =   0   'False
               Index           =   2
            End
         End
         Begin VB.Menu M1202020000 
            Caption         =   "Respuesta"
            Enabled         =   0   'False
            Index           =   0
            Begin VB.Menu M1202020100 
               Caption         =   "Rechazo de Solicitud"
               Enabled         =   0   'False
               Index           =   0
            End
            Begin VB.Menu M1202020100 
               Caption         =   "Aceptación de Solicitud"
               Enabled         =   0   'False
               Index           =   1
            End
            Begin VB.Menu M1202020100 
               Caption         =   "Reconsideración de Rechazo"
               Enabled         =   0   'False
               Index           =   2
            End
            Begin VB.Menu M1202020100 
               Caption         =   "Extorno"
               Enabled         =   0   'False
               Index           =   3
            End
         End
         Begin VB.Menu M1202030000 
            Caption         =   "Consulta Historial"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu M1203000000 
         Caption         =   "Configuración"
         Enabled         =   0   'False
         Index           =   0
         Begin VB.Menu M1203010000 
            Caption         =   "Parámetros de Solicitud de Cobertura"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu M1204000000 
         Caption         =   "Seg. Sepelio"
         Index           =   0
         Begin VB.Menu M1204010000 
            Caption         =   "Anulación"
            Index           =   0
         End
         Begin VB.Menu M1204010000 
            Caption         =   "Configuración"
            Enabled         =   0   'False
            Index           =   1
            Begin VB.Menu M1204010100 
               Caption         =   "Parámetros"
               Enabled         =   0   'False
               Index           =   0
            End
         End
         Begin VB.Menu M1204010000 
            Caption         =   "Solicitud de Activación"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu M1204010000 
            Caption         =   "Rechazo de Solicitud"
            Enabled         =   0   'False
            Index           =   3
         End
         Begin VB.Menu M1204010000 
            Caption         =   "Aceptación de Solicitud"
            Enabled         =   0   'False
            Index           =   4
         End
         Begin VB.Menu M1204010000 
            Caption         =   "Actualiza Datos"
            Enabled         =   0   'False
            Index           =   5
         End
      End
      Begin VB.Menu M1205000000 
         Caption         =   "Seg. MYPE"
         Enabled         =   0   'False
         Index           =   0
         Begin VB.Menu M1205010000 
            Caption         =   "Generación de Trama"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu M1205010000 
            Caption         =   "Activación"
            Enabled         =   0   'False
            Index           =   1
            Begin VB.Menu M1205010100 
               Caption         =   "Registro"
               Enabled         =   0   'False
               Index           =   0
            End
            Begin VB.Menu M1205010100 
               Caption         =   "Mantenimiento"
               Enabled         =   0   'False
               Index           =   1
            End
            Begin VB.Menu M1205010100 
               Caption         =   "Consulta"
               Enabled         =   0   'False
               Index           =   2
            End
         End
         Begin VB.Menu M1205010000 
            Caption         =   "Respuesta"
            Enabled         =   0   'False
            Index           =   2
            Begin VB.Menu M1205010200 
               Caption         =   "Rechazo Solicitud"
               Enabled         =   0   'False
               Index           =   0
            End
            Begin VB.Menu M1205010200 
               Caption         =   "Aceptación Solicitud"
               Enabled         =   0   'False
               Index           =   1
            End
            Begin VB.Menu M1205010200 
               Caption         =   "Reconsideración de Rechazo"
               Enabled         =   0   'False
               Index           =   2
            End
            Begin VB.Menu M1205010200 
               Caption         =   "Extorno"
               Enabled         =   0   'False
               Index           =   3
            End
         End
         Begin VB.Menu M1205010000 
            Caption         =   "Consulta Historial"
            Enabled         =   0   'False
            Index           =   3
         End
      End
      Begin VB.Menu M1206000000 
         Caption         =   "Gestión de Siniestros"
         Enabled         =   0   'False
         Index           =   0
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
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
' Esta sub se ejecuta al presionar las teclas
'****************APRI 20161024
Public Sub Ejecutar_HotKey(ByVal lParam As Long)
                
                
    Dim frm As String
    Dim nOperacion As Long
    Dim sDescOperacion As String

    If UCase(Screen.ActiveForm.Name) <> UCase("MDISicmact") And UCase(Screen.ActiveForm.Name) <> UCase("frmCajeroOperaciones") Then
        Exit Sub
    End If
    
    Select Case lParam
        Case 7340034 'CTRL+F1
            If gAtajoTeclado.bRenovaPigno = True Then
                nOperacion = gColPOpeRenovacEFE
                sDescOperacion = ""
                EjecutaOperacion nOperacion, sDescOperacion
            End If
        Case 7405570 'CTRL+F2
            frm = "frmCapOpePlazoFijo"
            If gAtajoTeclado.bRetiroIntDPF = True Then
                nOperacion = 210201
                sDescOperacion = "- RETIRO INTERES EFECTIVO"
                EjecutaOperacion nOperacion, sDescOperacion
            End If
        Case 7471106 'CTRL+F3
            If gAtajoTeclado.bHabEfectivo = True Then
                nOperacion = 901013
                sDescOperacion = "- TRANSFERENCIA DE EFECTIVO ENTRE CAJEROS"
                EjecutaOperacion nOperacion, sDescOperacion
            End If
        Case 7536642 'CTRL+F4
            If gAtajoTeclado.bDebitoSP = True Then
                nOperacion = gCapConSerPagDeb
                sDescOperacion = "- DEBITO SP"
                EjecutaOperacion nOperacion, sDescOperacion
            End If
        Case 7602178 'CTRL+F5
            If gAtajoTeclado.bCancelaDPF = True Then
                nOperacion = 210301
                sDescOperacion = "- CANCELACIÓN EFECTIVO"
                EjecutaOperacion nOperacion, sDescOperacion
            End If
        Case 7667714 'CTRL+F6
            If gAtajoTeclado.bCancelaPigno = True Then
                nOperacion = gColPOpeCancelacEFE
                sDescOperacion = "Cancelación Pignoraticio"
                EjecutaOperacion nOperacion, sDescOperacion
            End If
        Case 7733250 'CTRL+F7
            If gAtajoTeclado.bCompraME = True Then
                nOperacion = "900022"
                sDescOperacion = ""
                EjecutaOperacion nOperacion, sDescOperacion
            End If
        Case 7798786 'CTRL+F8
            If gAtajoTeclado.bRetiroEfectAho = True Then
                nOperacion = gAhoRetEfec
                sDescOperacion = "- RETIRO EFECTIVO"
                EjecutaOperacion nOperacion, sDescOperacion
            End If
        Case 7864322 'CTRL+F9
            If gAtajoTeclado.bDesembEfect = True Then
                nOperacion = gCredDesembEfec
                sDescOperacion = ""
                EjecutaOperacion nOperacion, sDescOperacion
            End If
        Case 7929858 'CTRL+F10
            If gAtajoTeclado.bConfirHab = True Then
                nOperacion = gOpeHabCajConfHabBovAge
                sDescOperacion = "CAJERO " & "- CONFIRMACIÓN DE HABILITACIÓN"
                EjecutaOperacion nOperacion, sDescOperacion
            End If
        Case 7995394 'CTRL+F11
            If gAtajoTeclado.bDesembCta = True Then
                nOperacion = gCredDesembCtaNueva
                sDescOperacion = ""
                EjecutaOperacion nOperacion, sDescOperacion
            End If
        Case 8060930 'CTRL+F12
            If gAtajoTeclado.bDepoEfectAHO = True Then
                nOperacion = 200201
                sDescOperacion = "- DEPÓSITO EFECTIVO"
                EjecutaOperacion nOperacion, sDescOperacion
            End If
        Case 3211265 'ALT+1
            If gAtajoTeclado.bVentaME = True Then
                nOperacion = gOpeCajeroMEVenta
                sDescOperacion = ""
                EjecutaOperacion nOperacion, sDescOperacion
            End If
        Case 3276801 'ALT+2
            If gAtajoTeclado.bAperGiro = True Then
                nOperacion = gServGiroApertEfec
                sDescOperacion = ""
                EjecutaOperacion nOperacion, sDescOperacion
            End If
        Case 3342337 'ALT+3
            If gAtajoTeclado.bRetiroCTS = True Then
                nOperacion = 220301
                sDescOperacion = "- RETIRO EFECTIVO"
                EjecutaOperacion nOperacion, sDescOperacion
            End If
        Case 3407873 'ALT+4
            If gAtajoTeclado.bPagoSERV = True Then
                nOperacion = gDepositoRecaudo
                sDescOperacion = ""
                EjecutaOperacion nOperacion, sDescOperacion
            End If
        Case 3473409 'ALT+5
            If gAtajoTeclado.bPagoCRED = True Then
                nOperacion = gCredPagNorNorEfec
                sDescOperacion = ""
                EjecutaOperacion nOperacion, sDescOperacion
            End If
        Case 3538945 'ALT+6
            If gAtajoTeclado.bCuadre = True Then
                nOperacion = gOpeHabCajRegEfect
                sDescOperacion = ""
                EjecutaOperacion nOperacion, sDescOperacion
            End If
        Case 3604481 'ALT+7
            If gAtajoTeclado.bAumentoCapDPF = True Then
                nOperacion = 210801
                sDescOperacion = "- AUMENTO CAPITAL PLAZO FIJO"
                EjecutaOperacion nOperacion, sDescOperacion
            End If
        Case 3670017 'ALT+8
            If gAtajoTeclado.bCancelGiro = True Then
                nOperacion = gServGiroCancEfec
                sDescOperacion = ""
                EjecutaOperacion nOperacion, sDescOperacion
            End If
        Case 3735553 'ALT+9
                frmPosicionCli.Show 1
    End Select
    
End Sub
'****************END APRI 20161024
Private Sub cmdVer_Click()
      'FrmCapAutOpeEstados.Show
End Sub

Private Sub Command1_Click()

Dim loPrevio As New previo.clsprevio
Dim lsCadImp As String
Dim i As Integer

i = 0

Do While MsgBox("Imprimir???", vbYesNo, "Aviso") = vbYes
    
    lsCadImp = ""
    lsCadImp = Chr$(27) & Chr$(64)
    lsCadImp = lsCadImp & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
    lsCadImp = lsCadImp & Chr$(27) & Chr$(15) 'Condensada
    lsCadImp = lsCadImp & Chr$(27) & Chr$(67) & Chr$(22) 'Longitud de página a 22 líneas'
    lsCadImp = lsCadImp & Chr$(27) & Chr$(77)  'Tamaño 10 cpi
    lsCadImp = lsCadImp & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
    
    lsCadImp = lsCadImp & Chr$(27) & Chr$(103)
    lsCadImp = lsCadImp & "   " & Chr(10)
    lsCadImp = lsCadImp & "   " & Chr(10)
    lsCadImp = lsCadImp & Chr$(27) & Chr$(77)
    
    lsCadImp = lsCadImp & Chr$(27) & Chr$(69) 'activa negrita
    lsCadImp = lsCadImp & Chr$(27) + Chr$(72) ' desactiva negrita
     
    If i > 0 Then
        lsCadImp = lsCadImp & "" & Chr(10)
        lsCadImp = lsCadImp & "" & Chr(10)
        lsCadImp = lsCadImp & "" & Chr(10)
        lsCadImp = lsCadImp & "" & Chr(10)
    End If
     
     
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & space(28) & "89123451231231321313212316789" & space(26) & "1234567890" & space(38) & "3456789123123123132" & Chr(10)
    lsCadImp = lsCadImp & space(83) & "1234567890" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & space(38) & "Apellidos" & Chr(10)
    lsCadImp = lsCadImp & space(38) & "Nombre" & Chr(10)
    lsCadImp = lsCadImp & space(38) & "DNI123456789123456789123456789123456789123456789123456789213456789123456789" & space(24) & "9999" & Chr(10)
    lsCadImp = lsCadImp & space(38) & "Calle" & Chr(10)
    lsCadImp = lsCadImp & space(38) & "Urbanizacion" & Chr(10)
    lsCadImp = lsCadImp & space(38) & "Ciudad" & Chr(10)
    lsCadImp = lsCadImp & space(38) & "Telefono" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & space(38 + 79) & "3456789123456789123 999" & Chr(10)
    lsCadImp = lsCadImp & space(38 + 79) & "3456789123456789123 999" & Chr(10)
    lsCadImp = lsCadImp & space(38 + 79) & "3456789123456789123 999" & Chr(10)
    lsCadImp = lsCadImp & space(38 + 79) & "3456789123456789123 999" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & space(22) & "3456789123456789123 999" & Chr(10)
    lsCadImp = lsCadImp & space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & space(22) & "3456789123456789123 999" & Chr(10)
    lsCadImp = lsCadImp & space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lsCadImp = lsCadImp & space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & space(22) & "3456789123456789123 999" & Chr(10)
    lsCadImp = lsCadImp & space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & space(22) & "3456789123456789123 999" & Chr(10)
    lsCadImp = lsCadImp & space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & space(22) & "3456789123456789123 999" & Chr(10)
    lsCadImp = lsCadImp & space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lsCadImp = lsCadImp & space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lsCadImp = lsCadImp & space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lsCadImp = lsCadImp & space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & space(22) & "3456789123456789123 999" & Chr(10)
    lsCadImp = lsCadImp & space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lsCadImp = lsCadImp & space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lsCadImp = lsCadImp & space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lsCadImp = lsCadImp & space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lsCadImp = lsCadImp & space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lsCadImp = lsCadImp & space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lsCadImp = lsCadImp & space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lsCadImp = lsCadImp & space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lsCadImp = lsCadImp & space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lsCadImp = lsCadImp & space(20) & "123456712345678912345678912345678912345678912345678912345678912345678912345" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & space(30) & "Inc. 11: pasados 30 dias del vencimiento de su contrato" & Chr(10)
    lsCadImp = lsCadImp & space(30) & "sus joyas entrarán a Remate Sin Notificar" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & "" & Chr(10)
    lsCadImp = lsCadImp & space(116) & "123456789123456789123456789" & Chr(10)
    lsCadImp = lsCadImp & space(20) & "12345678912345678912345678912345678912345678912345678912345678912345678912" & space(2) & "5678912345" & space(19) & "123456789" & Chr(10)
    lsCadImp = lsCadImp & Chr$(27) + Chr$(18) ' cancela condensada

    loPrevio.PrintSpool sLpt, lsCadImp, False
    
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

'Private Sub M0201000000_Click(Index As Integer)
'    Select Case Index
'        Case 4
'            frmCapSimulacionPF.Show 1
'
'        Case 9
'           frmCapReportes.Show 1
'           ' frmRepConsolidados.Show
'
'        Case 10
'            'frmAclAhorros.Show
'
'        Case 11
'            'frmAclColocaciones.Show
'
'        Case 15
'            frmExoneracionITF.Show
'
'        Case 16
'            frmCapNoCobroInactivas.Show 1
'        Case 20
'            frmCapBloqueoDesbloqueoParcial.Inicia gCapAhorros
'        Case 21
'            frmCapPlazoFijoBloqueo.Show 1
'        Case 22
'            FrmCompraVentaAut.Show 1
'        Case 25
'            'frmCapReasigGest.Show 1
'        'MIOL RQ12257 ***************************************
'        Case 31
'            frmReqSunat.Show 1
'        'END MIOL *******************************************
'         'MIOL RQ12257 ***************************************
'        'Case 32
'         '   frmNivelesAprobacionCVAutorizar.InicioAutorizarNiveles
'        'END MIOL *******************************************
'    End Select
'End Sub

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
    Case 0
        'frmPreparaSorteo.Inicia ("00")
        
    Case 1
        'frmConsolidaSorteo.Inicia ("00")
  End Select
End Sub

Private Sub M0201080200_Click(Index As Integer)
Select Case Index
    Case 0
        'frmPreparaSorteo.Inicia (gsCodAge)
        
    Case 1
        'frmConsolidaSorteo.Inicia (gsCodAge)
  End Select
End Sub

'JUEZ 20130828 **********************************
'Private Sub M0201010100_Click(Index As Integer)
'    Select Case Index
'        Case 0
'             frmCapParametros.Inicia False
'        Case 1
'            frmCapParametros.Inicia True
'    End Select
'End Sub

'Private Sub M0201010200_Click(Index As Integer)
'    Select Case Index
'        Case 0
'             frmCapParametrosCom.Inicia 1 'Registro
'        Case 1
'            frmCapParametrosCom.Inicia 2 'Consulta
'        Case 2
'            frmCapParametrosCom.Inicia 3 'Mantenimiento
'    End Select
'End Sub
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

''RECO20140607 ERS008-2014***************************************************
'Private Sub M0201010400_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Mantenimiento
'            frmGiroTarifarioMant.Inicio 1, "Giros - Tarifario - Mantenimiento"
'        Case 1 'consulta
'            frmGiroTarifarioMant.Inicio 2, "Giros - Tarifario - Consulta"
'    End Select
'End Sub
''RECO FIN*******************************************************************

Private Sub M0201010500_Click(Index As Integer)
Select Case Index
    'Case 0 'Mantenimiento
     '       frmNivelesAprobacionCV.InicioRegistroNiveles
    'Case 1 'Consulta
     '       frmNivelesAprobacionCVConsulta.InicioConsultaNiveles
End Select
End Sub

''JUEZ 20140908 ***********************************************/
'Private Sub M0201010601_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Mantenimiento
'            frmCapParametros_NEW.Mantenimiento gCapAhorros
'        Case 1 'Consulta
'            frmCapParametros_NEW.Consulta gCapAhorros
'    End Select
'End Sub
'
'Private Sub M0201010602_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Mantenimiento
'            frmCapParametros_NEW.Mantenimiento gCapPlazoFijo
'        Case 1 'Consulta
'            frmCapParametros_NEW.Consulta gCapPlazoFijo
'    End Select
'End Sub

'Private Sub M0201010603_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Mantenimiento
'            frmCapParametros_NEW.Mantenimiento gCapCTS
'        Case 1 'Consulta
'            frmCapParametros_NEW.Consulta gCapCTS
'    End Select
'End Sub
''END JUEZ ****************************************************/

'Private Sub M0201020201_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Mantenimiento
'            frmCapTasaInt.Inicia gCapPlazoFijo, False
'        Case 1 'Mantenimiento
'            'frmCapTasaIntPF.Inicia gCapPlazoFijo, False
'            frmCapTasaInt.Inicia gCapPlazoFijo, False, True 'JUEZ 20140220
'    End Select
'End Sub

'Private Sub M0201020202_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Consulta
'            frmCapTasaInt.Inicia gCapPlazoFijo, True
'        Case 1 'Consulta
'            'frmCapTasaIntPF.Inicia gCapPlazoFijo, True
'            frmCapTasaInt.Inicia gCapPlazoFijo, True, True 'JUEZ 20140220
'    End Select
'End Sub


'Private Sub M0201100000_Click(Index As Integer)
'
'    Select Case Index
'        Case 0 'Niveles
'            FrmCapRegNivelAutRetCan.Show 1
'        Case 1 'Niveles-Grupos
'            FrmCapRegNivelAutRetCanDet.Show 1
'        Case 2 'Aprobacion / rechazo
'            'If Not VerificarRFIII Then Exit Sub ' *** RIRO SEGUN TI-ERS108-2013 ***
'            FrmCapRegAproAutRetCan.Show 1
'    End Select
'
'End Sub

'Private Sub M0201020100_Click(Index As Integer)
'    Select Case Index
'        Case 0 'mantenimiento
'            frmCapTasaInt.Inicia gCapAhorros, False
'        Case 1 'Consulta
'            frmCapTasaInt.Inicia gCapAhorros, True
'    End Select
'End Sub

'Private Sub M0201020200_Click(Index As Integer)
'    Select Case Index
''        Case 0 'Mantenimiento
''            frmCapTasaInt.Inicia gCapPlazoFijo, False
''        Case 1 'Consulta
''            frmCapTasaInt.Inicia gCapPlazoFijo, True
'        Case 2 'Cambio de Tasa Plazo Fijo
'            frmCapCambioTasa.Inicia gCapPlazoFijo
'        'By Capi 07082008
'        Case 3 'Cambio de Tasa Plazo Fijo
'            frmCapCambioTasa.Inicia gCapPlazoFijo, False
'        '
'    End Select
'End Sub
'
'Private Sub M0201020300_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmCapTasaInt.Inicia gCapCTS, False
'        Case 1
'            frmCapTasaInt.Inicia gCapCTS, True
'        Case 2 'JUEZ 20140228
'            frmCapCambioTasaCTSLote.Show 1
'    End Select
'End Sub

Private Sub M0201030000_Click(Index As Integer)
    'Mantenimiento
    Select Case Index
        Case 0 'Ahorro
            frmCapMantenimiento.Inicia gCapAhorros
        Case 1 'Plazo fijo
            frmCapMantenimiento.Inicia gCapPlazoFijo
        Case 2 'Cts
            frmCapMantenimiento.Inicia gCapCTS
    End Select
End Sub
'
'Private Sub M0201040000_Click(Index As Integer)
'    'Bloqueos / Desbloqueos
'    Select Case Index
'        Case 0 'Ahorros
'            'If Not VerificarRFIII Then Exit Sub ' *** RIRO SEGUN TI-ERS108-2013 ***
'            frmCapBloqueoDesbloqueo.Inicia gCapAhorros
'        Case 1 'Plazo Fijo
'            'If Not VerificarRFIII Then Exit Sub ' *** RIRO SEGUN TI-ERS108-2013 ***
'            frmCapBloqueoDesbloqueo.Inicia gCapPlazoFijo
'        Case 2 'Cts
'            'If Not VerificarRFIII Then Exit Sub ' *** RIRO SEGUN TI-ERS108-2013 ***
'            frmCapBloqueoDesbloqueo.Inicia gCapCTS
'    End Select
'End Sub

'Private Sub M0201050000_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Registro
'            frmTarjetaRegistra.Show 1
'        Case 1 'Relacion
'            frmTarjetaBloqueo.Show 1
'        Case 2 'Cambio de Clave
'            FrmTarjetaCambioClave.Show 1
'    End Select
'End Sub

'Private Sub M0201060000_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmCapBeneficiario.Inicia False
'        Case 1
'            frmCapBeneficiario.Inicia True
'    End Select
'End Sub

'Private Sub M0201070000_Click(Index As Integer)
'    Select Case Index
'        'Case 0 'generacion
'            'frmCapOrdPagGenEmi.Show 1
'        Case 1 'Certificacion
'            frmCapOrdPagAnulCert.Inicia gAhoOPCertificacion
'        Case 2 'Anulacion
'            frmCapOrdPagAnulCert.Inicia gAhoOPAnulacion
'        Case 3 'Consulta
'            frmCapOrdPagConsulta.Show 1
'        Case 4
'            frmIngOP.Show 1
'        Case 5
'            'frmCartasOrdenes.Show 1
'        'MIOL 20121006, SEGUN RQ12272 **********
'        Case 6
'            frmDesbloqueoClienteOrdPago.Show
'        'END MIOL ******************************
'    End Select
'End Sub

'Private Sub M0201070100_Click(Index As Integer)
'Select Case Index
'    Case 0
'        frmCapOrdPagSolicitud.Show 1
'    Case 1
'        frmCapOrdPagEmiteImpr.Show 1
''        frmCapOrdPagProceso.Inicia gCapTalOrdPagEstSolicitado
''    Case 2
''        frmCapOrdPagProceso.Inicia gCapTalOrdPagEstEnviado
''    Case 3
''        frmCapOrdPagProceso.Inicia gCapTalOrdPagEstRecepcionado
'End Select
'End Sub



 

'Private Sub M0201090000_Click(Index As Integer)
'Select Case Index
'    Case 0
'        frmCapPersParam.Inicia gCapAhorros
'    Case 1
'        frmCapPersParam.Inicia gCapPlazoFijo
'    Case 2
'        frmCapPersParam.Inicia gCapCTS
'End Select
'End Sub
'
'Private Sub M0201110000_Click(Index As Integer)
'Select Case Index
'    Case 0
'        frmCapTasaEspSeg.Show 1
'    Case 1
'        frmCapTasaEspAprRech.Show 1
'    Case 2
'        frmExtornoSolTasaEspecial.Show 1
'    Case 4
'        frmCapAdmNiveles.Show 1 ' Agregado Por RIRO el 20130418
'
'End Select
'End Sub

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
'
'Private Sub M0201130000_Click(Index As Integer)
'    Select Case Index
'    Case 0 'Renovacion de Plazo
'        FrmCapPlazoPanderito.Show 1
'    End Select
'End Sub
'MADM 20110318 - SERVICIOS
'Private Sub M0201140000_Click(Index As Integer)
'     Select Case Index
'        Case 0
'            'frmCredGeneraComisionPagoVarios.Show 1
'        Case 1 'Gravar Garantias
'            'frmCredGeneraLecturaFilePagoVarios.Show
'        Case 2 'Solicitud CartaFianza
'            'frmCredGeneraReportePagoVarios.Show 1
'
'        'Agregado por RIRO el 20130418 ****
'        Case 4
'            frmCapCargaArchivo.Show 1
'
'        Case 5
'            frmCapGeneraArchivoRecaudo.Show 1
'        ' Fin RIRO ****
'
'    End Select
'End Sub
''END MADM

'' Agregado por RIRO el 20130418 *****
'Private Sub M0201140100_Click(Index As Integer)
'
'    Select Case Index
'
'        Case 0
'                frmCapRegistroConvenioAhorros.Inicia 1
'
'        Case 1
'                frmCapRegistroConvenioAhorros.Inicia 2
'
'        Case 2
'                frmCapRegistroConvenioAhorros.Inicia 3
'
'    End Select
'
'End Sub
'' Fin RIRO ****


'***Agregado por ELRO el 20130716, según RFC1306270002****
'Private Sub M0201140200_Click(Index As Integer)
'Select Case Index
'    Case 0 'Registro de Convenio de Servicio de Pago
'        frmCapServicioPago.iniciarRegistro
'    Case 1 'Mantenimiento de Convenio de Servicio de Pago
'        frmCapServicioPago.inicarMantenimiento
'    Case 2 'Cargar Trama de Convenio de Servicio de Pago
'        frmCapServicioPagoCargaArchivo.Show 1
'    Case 3 'Baja Trama de Convenio de Servicio de Pago
'        frmCapServicioPagoBajaArchivo.Show 1
'End Select
'End Sub
'***Fin Agregado por ELRO el 20130716, según RFC1306270002
'
''BRGO 20110425 - REGISTRO DE SUELDOS DE CLIENTES CTS
'Private Sub M0201150000_Click(Index As Integer)
'     Select Case Index
'        Case 0 'Registro individual de sueldos CTS
'            frmCapDatosSueldosCTS.Show 1
'        Case 1 'Registro en lote de sueldos CTS
'            frmCapCargaArchivoCTS.Show 1
'        Case 2 'JUEZ 20140305
'            frmCapCambioEstadoCTS.Show 1
'    End Select
'End Sub
''END BRGO
'
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

'Private Sub M0201170000_Click(Index As Integer)
'    Select Case Index
'    Case 0
'        frmParametrosPIT.Inicia False
'    Case 1
'        frmParametrosPIT.Inicia True
'    End Select
'End Sub
'***Agregado por ELRO el 20120905, según OYP-RFC087-2012
'Private Sub M0201180100_Click(Index As Integer)
'Select Case Index
'    Case 0: frmCapIndConCli.Show 1
'    Case 1: frmCapIndConCliVoBo.Show 1
'End Select
'End Sub

Private Sub M0201180200_Click(Index As Integer)
    'frmCapIndComDep.Show 1
End Sub
'***Fin Agregado por ELRO el 20120905******************

'Private Sub M0301010000_Click(Index As Integer)
'    Select Case Index
'    Case 0 'Solicitud CartaFianza
'        frmCFSolicitud.Show 1
'    Case 1 'Gravar Garantias
'        frmCredGarantCred.Inicio PorMenu, , 1
'    Case 2 'Solicitud CartaFianza
'        Call frmCFSugerencia.Inicia
'    Case 4 'Emitir CartaFianza
'        frmCFEmision.Show 1
'    Case 5 ' Honrar CartaFianza
'        FrmCFHonrar.Show 1
'    'Case 8 ' Niveles de Aprobacion
'        'frmCFNivelApr.Show 1 YA NO SE USA
'
'    Case 8 'Matenimiento de Tarifario
'        'frmCFTarifario.Show 1
'    Case 9  ' Relacionar con Credito
'        'FrmCFHonrarCredito.Show 1
'    'WIOR 20130312 COMENTO INICIO ********************************************
'    'Case 10  ' Relacionar con Credito
'    '    FrmCFRenovacion.Show 1
'    'WIOR 20130312 COMENTO FIN ************************************************
'    'WIOR 20120613 **************************************************************
'    Case 11  'Dar de Baja folio Hoja CF
'        frmCFDarBaja.Show 1
'    'WIOR FIN *******************************************************************
'    'WIOR 20121015 **************************************************************
'    Case 12  'Mantenimiento de Modalidad
'        frmCFMantModalidad.Show 1
'    'WIOR FIN *******************************************************************
'    End Select
'End Sub

'Private Sub M0301010100_Click(Index As Integer)
'    Select Case Index
'    Case 0 'Aprobacion
'         frmCFAprobacion.Show 1
'    Case 1 ' 'Rechazo
'        FrmCFRechazar.Show 1
'    Case 2
'        FrmCFRetirarApr.Show 1
'    Case 3
'        FrmCFDevolucion.Show 1
'    'By capi Set07
'    Case 4
'        FrmCFCancelacion.Show 1
'   'MADM 20121201
'    Case 5
'        frmCFExtornoApro.Show 1
'    'WIOR 20120828
'    Case 6
'        frmCFModEmision.Show 1
'    'WIOR 20130311 ***************************
'    Case 7: frmCFEditarModEm.Show 1
'    'WIOR FIN ********************************
'    End Select
'
'End Sub

'Private Sub M0301010200_Click(Index As Integer)
'    Select Case Index
'    Case 0 'Consultas
'        frmCFHistorial.Show 1
'
'    End Select
'End Sub

'Private Sub M0301010300_Click(Index As Integer)
'    Select Case Index
'    Case 0 'Reportes
'        frmCFReporte.Show 1
'    End Select
'
'End Sub
''WIOR 20130312 ***************************************
'Private Sub M0301010600_Click(Index As Integer)
' Select Case Index
'    Case 0: frmCFAutRenovacion.Show 1
'    Case 1: frmCFExtornoRenovacion.Inicio (1)
'    Case 2: FrmCFRenovacion.Show 1
'    Case 3: frmCFExtornoRenovacion.Inicio (2)
' End Select
'End Sub
'WIOR FIN *******************************************
'Private Sub M0301020000_Click(Index As Integer)
'    Select Case Index
'    Case 8 'Juez 20120905 '7 'Refinanciacion de Credito
'        Call frmCredSolicitud.RefinanciaCredito(Registrar)
'    Case 9 'Juez 20120905 '8 'Actualizacion con Metodos de Liquidacion
'        frmCredMntMetLiquid.Show 1
'    Case 10 'Juez 20120905 '9 'Perdonar Mora
'        frmCredPerdonarMora.Show 1
'    Case 12 'Juez 20120905 '11 'Reasignar Institucion
'        frmCredReasigInst.Show 1
'    Case 13 'Juez 20120905 '12 'Transferencia a Recuperaciones
'        frmCredTransARecup.Show 1
'    Case 18 'Juez 20120905 '17 'Registro de Dacion de Pago
'        'frmCredRegisDacion.Show 1
'    Case 19 'Juez 20120905 '18
'        'frmCredCargoAuto.Show 1
'    Case 20 'Juez 20120905 '19
'        frmCredCodModular.Show 1
'    Case 21 'Juez 20120905 '20
'        frmCredAsigCComodin.Show 1
'    Case 22 'Juez 20120905 '21
'        'frmCredAdmPrepago.Show 1
'    Case 23
'        frmCredValorizaCheque.Show 1
'    ' CMACICA_CSTS - 05/11/2003 -------------------------------------------------
'    Case 24
'        'frmCredCalendarioDesemb.Show 1
'    Case 25
'        Call frmCredSolicitud.SustitucionCredito(Registrar)
'    ' --------------------------------------------------------------------------
'    'ALPA 20091007***********************************************
''    Case 36
'         'Call frmCredSolicitud.AmpliacionCredito(Registrar)
'    '************************************************************
'    Case 27
'        frmCredConvRegDev.Show 1
'    Case 28
'         'FrmVerRFA.Show vbModal
'
'    Case 29
'         'frmCredAutorizar.Show 1
'    Case 30
'        'frmCredCalendCOFIDE.Show 1
'    Case 31
'        'FrmCredCambioLC.Show vbModal
'    Case 32
'        'ARCV 14-02-2007
'        Call frmCredSolicitud.AmpliacionCredito(Registrar)
'        'frmCredEduacionGeneraTxt.Show 1
'
''        Dim sCad As String
''        Dim sArchivo As String
''        Dim NumeroArchivo As Integer
''        Dim oCred As COMNCredito.NCOMCredDoc
''        Set oCred = New COMNCredito.NCOMCredDoc
''
''        sArchivo = App.path & "\Spooler\" & "A" & Mid(CStr(Year(gdFecSis)), 3, 2) & Format(CStr(Month(gdFecSis)), "00") & "31.151"
''        sCad = oCred.GeneraArchivoTXT_Educacion(CStr(Year(gdFecSis)) & Format(CStr(Month(gdFecSis)), "00"), "A")
''
''        NumeroArchivo = FreeFile
''        Open sArchivo For Output As #NumeroArchivo
''        Print #NumeroArchivo, sCad
''        Close #NumeroArchivo
''        MsgBox "Se ha generado el Archivo " & "A" & Mid(CStr(Year(gdFecSis)), 3, 2) & Format(CStr(Month(gdFecSis)), "00") & "31.151" & " Satisfactoriamente", vbInformation, "Mensaje"
''
''        sArchivo = App.path & "\Spooler\" & "C" & Mid(CStr(Year(gdFecSis)), 3, 2) & Format(CStr(Month(gdFecSis)), "00") & "31.151"
''
''        sCad = oCred.GeneraArchivoTXT_Educacion(CStr(Year(gdFecSis)) & Format(CStr(Month(gdFecSis)), "00"), "C")
''
''        NumeroArchivo = FreeFile
''        Open sArchivo For Output As #NumeroArchivo
''        Print #NumeroArchivo, sCad
''        Close #NumeroArchivo
''        MsgBox "Se ha generado el Archivo " & "C" & Mid(CStr(Year(gdFecSis)), 3, 2) & Format(CStr(Month(gdFecSis)), "00") & "31.151" & " Satisfactoriamente", vbInformation, "Mensaje"
''
''        Set oCred = Nothing
'         Case 35
'            FrmCredGestionCobranza.Show 1
'        'ALPA 20091007*****************************************
'         Case 36
'            frmGruposEconomicos.Show 1
'        '******************************************************
'        'JACA 20110628*****************************************
''        Case 39
''            frmCredRegistrarParametroBPPR.Show 1
'        'JACA END******************************************
'
'        Case 40 'JACA 20120109
'            'frmMetasAgencias.Show 1
'        'WIOR 20130411 COMENTO LINEAS ABAJO******************************
'        'Case 44 'WIOR 20120828
'        '    frmCredRiesgos.Show 1
'        'WIOR 20131017 **************************************************
'         Case 47
'            'FrmCredTraspCartera.Show 1
'        'WIOR FIN *******************************************************
'    End Select
'End Sub


'Private Sub M0301020100_Click(Index As Integer)
'        Select Case Index
'            Case 10 'Registra Configuracion de Clientes Preferenciales
'                frmCredConfClientesPreferenciales.Registrar
'    End Select
'End Sub

'Private Sub M0301020101_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Mantenimeinto de Parametros
'            frmCredMantParametros.InicioActualizar
'        Case 1 'Consulta de Parametros
'            frmCredMantParametros.InicioCosultar
'    End Select
'End Sub

'Private Sub M0301020102_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Registro de Lineas de Credito
'            'frmCredLineaCredito.Registrar
'        Case 1 'Mantenimiento de Lineas de Credito
'            'frmCredLineaCredito.actualizar
'        Case 2 ' Consulta de lineas de Credito
'            'frmCredLineaCredito.Consultar
'    End Select
'End Sub
'
'Private Sub M0301020103_Click(Index As Integer)
''    Select Case Index
''        Case 0 'Mantenimeinto de Niveles de Aprobacion
''            'frmCredNivAprCred.inicio MuestraNivelesActualizar
''            frmCredNivAprobacion.Show 1 '08-06-2006
''        Case 1 'Consulta de Niveles de Aprobacion
''            frmCredNivAprCred.Inicio MuestraNivelesConsulta
''    End Select
'
'    'JUEZ 20121204 **********************************************
'    'MAVM 20110523 ***
'    'Select Case Index
'    '    Case 0
'    '        frmCredConfNivelAprobacion.Show vbModal
'    '    Case 1
'    '    frmCredConfNivelAprobacionListar.Show vbModal
'    'End Select
'    '***
'    Select Case Index
'        Case 0
'            frmCredNewNivAprGrupoApr.InicioGrupoAprobacion
'        Case 1
'            frmCredNewNivAprGrupoApr.InicioRegistroNiveles
'        Case 2
'            frmCredNewNivAprParamApr.Show 1
'        Case 3
'            frmCredNewNivAprDelegacion.Inicio
'    'End Select
'    'END JUEZ ****************************************************
'    'ALPA 20150215**************************
'    Case 4
'            frmParCargosTasa.Show 1
'    End Select
'    '***************************************
'End Sub

'Private Sub M0301020104_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Mantenimeinto de Gastos
'            frmCredMntGastos.Inicio InicioGastosActualizar
'        Case 1
'            frmCredMntGastos.Inicio InicioGastosConsultar
'    End Select
'End Sub
'
'Private Sub M0301020105_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            FrmCampanas.Registro
'        Case 1
'            FrmCampanas.Mantenimiento
'        Case 2
'            FrmCampanas.Consultas
'         Case 3
'            frmCasaComercial.Show 'MADM 20100726
'    End Select
'
'End Sub

'Private Sub M0301020106_Click(Index As Integer)
'    Select Case Index
'        Case 0 'JUEZ 20120905
'            frmCredEvalParamTipos.Show 1
'        Case 1 'JUEZ 20120905
'            frmCredEvalParamIndicador.Show 1
'        Case 2 'WIOR 20120905
'            frmCredEvalParamEspecializacion.Show 1
'    End Select
'End Sub

'' JUEZ 20121204 **********************************
'Private Sub M0301020107_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmCredNewNivAprExoneracion.TiposExoneracion
'        Case 1
'            frmCredNewNivAprExoneracion.NivelesExoneracion
'        End Select
'End Sub
' END JUEZ ***************************************

Private Sub M0301020108_Click(Index As Integer)
    Select Case Index
        Case 0 'Mantenimeinto
            'frmCredAdmPeriodo.InicioActualizar
        Case 1 'Consulta
            'frmCredAdmPeriodo.InicioConsultar
    End Select
End Sub
''WIOR 20131123 **************************************
'Private Sub M0301020109_Click(Index As Integer)
'    'ALPA 20150215****************
'    Select Case Index
'        Case 0
'    '********************************
'            frmCredConfigCuotaBalon.Show 1
'    'ALPA 20150215****************
'        Case 1
'            FrmCredLineaCreditoConfiguracion.Show 1
'    End Select
'    '********************************
'End Sub
''WIOR FIN *******************************************

'JUEZ 20140530 *************************************
'Private Sub M0301020111_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Mantenimiento
'            frmCredMantLimSecEcon.Inicia 1
'        Case 1 'Consulta
'            frmCredMantLimSecEcon.Inicia 2
'    End Select
'End Sub
'END JUEZ ******************************************

'Private Sub M0301020200_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Registro de Solicitud
'            frmCredSolicitud.Inicio Registrar
'        Case 1 'Consulta de Solicitud
'            frmCredSolicitud.Inicio Consulta
'    End Select
'End Sub

'Private Sub M0301020300_Click(Index As Integer)
'    Dim oCredRel As New UCredRelac_Cli  'COMDCredito.UCOMCredRela   'UCredRelacion
'
'    Select Case Index
'        Case 0 'Mantenimiento de Relaciones de Credito
'            frmCredRelaCta.Inicio oCredRel, InicioMantenimiento, InicioRegistroForm
'        '    Set oCredRel = Nothing
'        Case 1 'Reasignacion de Cartera en Lote
'            frmCredReasigCartera.Show 1
'        Case 2 'Consulta de Relaciones de Credito
'            frmCredRelaCta.Inicio oCredRel, InicioMantenimiento, InicioConsultaForm
'        '    Set oCredRel = Nothing
'        Case 3 'Confirmacion de reasignacion de cartera
'            'frmCredRelaCta.Inicio oCredRel, InicioMantenimiento, InicioConsultaForm
'            'frmCredConfirmaReasignaCartera.Show 1 'FRHU2014020 RQ14011
'            frmCredConfirmaReasignaCartera.Inicio 'FRHU2014020 RQ14011
'        Case 4 'FRHU 20140220 RQ14010 Asignacion de Agencia - Jefe de Negocios Territoriales
'            frmAsigAgeJNTerritorial.Inicio
'    End Select
'    Set oCredRel = Nothing
'End Sub

'Private Sub M0301020400_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Registro de Garantia
'            If gsProyectoActual = "H" Then
'                frmPersGarantiasHC.Inicio RegistroGarantia
'            Else
'                frmPersGarantias.Inicio RegistroGarantia
'            End If
'        Case 1 'Mantenimiento de Garantia
'            If gsProyectoActual = "H" Then
'                frmPersGarantiasHC.Inicio MantenimientoGarantia
'            Else
'                frmPersGarantias.Inicio MantenimientoGarantia
'            End If
'        Case 2 'Consulta de Garantia
'            If gsProyectoActual = "H" Then
'                frmPersGarantiasHC.Inicio ConsultaGarant
'            Else
'                frmPersGarantias.Inicio ConsultaGarant
'            End If
'        Case 3 'Gravament
'            frmCredGarantCred.Inicio PorMenu
'        Case 4 'Liberar Garantia
'            frmCredLiberaGarantia.Show 1
'        Case 5
'            'FrmMantGarantias.Show vbModal
'            frmGarantiaConf.Show 1
'        Case 6
'           'FrmCredRelGarantias.Show vbModal
'        Case 7
'           FrmGraAmpliado.Show vbModal
'        'MAVM 20100723 ***
'        Case 8
'           frmJoyGarRegistro.Show vbModal
'        Case 9
'           frmJoyGarRescate.Show vbModal
'        '***
'        'madm 20110420 ------------------
'        Case 10
'           frmCredGarantRealLegal.Show vbModal
'      '----------------------------------
'       'madm 20110525 ------------------
'        Case 11
'           frmCredGarantiaVerificaLegal.Show vbModal
'      '----------------------------------
'      'ALPA 20140318***********************************
'        Case 12
'           frmCredGarantiaReemplazo.Show vbModal
'      '************************************************
'    End Select
'End Sub

'Private Sub M0301020500_Click(Index As Integer)
'    'JUEZ 20121219 ******************
'    Select Case Index
'        Case 0 'Registro de Sugerencia
'            If gnAgenciaCredEval = 0 Then
'                frmCredSugerencia.Sugerencia lSugerTipoActRegistro
'            Else
'                frmCredSugerencia_NEW.Sugerencia lSugerTipoActRegistro
'            End If
'        Case 1 'Consulta de Sugerencia
'            If gnAgenciaCredEval = 0 Then
'                frmCredSugerencia.Sugerencia lSugerTipoActConsultar
'            Else
'                frmCredSugerencia_NEW.Sugerencia lSugerTipoActConsultar
'            End If
'    End Select
'    'END JUEZ ***********************
'End Sub

'Private Sub M0301020600_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Aprobacion de Credito
'            frmCredAprobacion.Show 1
'        Case 1 'Rechazo de Credito
'            frmCredRechazo.Rechazar
'        Case 2 'Anulacion de Credito
'            frmCredRechazo.Retirar
'        Case 3
'            FrmExtornoAprobacion.Show vbModal
'        Case 4
'            frmCredDesBloqCred.Show vbModal
'        Case 5
'            '20060329
'            'En esta línea se rechaza la solictud de credito
'            frmCredRechazo.Rechazar 3
'        Case 6
'            '20060329
'            'en esta línea se rechaza la sugerencia de credito
'            frmCredRechazo.Rechazar 4
'        'JUEZ 20121204 ********************
'        'MAVM 20110523 ***
'        Case 7
'            'frmCredPreAprobacion.Show vbModal
'            frmCredNewNivAprPorNivel.Inicio
'
'        Case 8
'            'frmCredPreAprobacionListar.Show vbModal
'            frmCredNewNivAprHist.Inicio
'        '***
'        'END JUEZ *************************
'        'ALPA 20150215*********************
'        Case 9
'            frmCredCargosTasa.Show 1
'        '**********************************
'
'    End Select
'End Sub

'Private Sub M0301020700_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Reprogramacion de Credito
'            frmCredReprogCred.Show 1
'        Case 1 'Reprogramacion en Lote
'            frmCredReprogLote.Show 1
'        Case 2
'            'frmReestructuraRFA.Show 1
'            frmCredReprogCredConvenio.Show 1 'WIOR 20140526
'    End Select
'End Sub

'Private Sub M0301020800_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Administracion de Gastos en Lote
'            'frmCredAsigGastosLote.Show 1
'        Case 1 ' mantenimiento de Penalidad
'            frmCredExonerarPen.Show 1
'        Case 2
'            frmCredAdmiGastos.Show 1
'        Case 3
'            'frmColAsignacionGastoLote.Show 1
'    End Select
'End Sub

'Private Sub M0301020900_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Nota del Analista
'            frmCredAsigNota.Show 1
'        Case 1 'Meta del Analista
'            frmCredMetasAnalista.Show 1
'    End Select
'End Sub

'Private Sub M0301021000_Click(Index As Integer)
'    Dim MatCalend As Variant
'    Dim Matriz(0) As String
'
'    Select Case Index
'        Case 0 'Calendario de Pagos
'            frmCredCalendPagos.Simulacion DesembolsoTotal
'        Case 1 'Desembolsos Parciales
'            frmCredCalendPagos.Simulacion DesembolsoParcial
'        Case 2 'Cuota Libre
''            MatCalend = frmCredCalendCuotaLibre.CalendarioLibre(True, gdFecSis, Matriz, 0#, 0, 0#)
'        Case 3
'            frmCredSimuladorPagos.Show 1
'        Case 4
'            'frmCredSimNroCuotas.Show 1
'            frmCredSimuladorGarantiaPlazo.Show 1 'JUEZ 20140226
'    End Select
'End Sub

'Private Sub M0301021100_Click(Index As Integer)
'    If Index = 0 Then
'        frmCredConsulta.Show 1
'    Else
'        'frmCredHistCalendario.Show 1
'        Call frmCredHistCalendario.Inicio
'    End If
'End Sub
'
'Private Sub M0301021200_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmCredDupDoc.Show 1
'        Case 1
'            frmCredReportes.Inicia "Reportes de Créditos"
'        Case 2
'            frmCredVinculados.Ini True, "Créditos Vinculados - Titulares"
'        Case 3
'            frmCredVinculados.Ini False, "Créditos Vinculados - T y G Consolidado"
'    End Select
'End Sub



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

'Private Sub M0301021500_Click(Index As Integer)
'Select Case Index
'    Case 1
'        Dim oGen As COMDConstSistema.DCOMGeneral   'DGeneral
'        Dim lbCierreRealizado As Boolean
'
'        Set oGen = New COMDConstSistema.DCOMGeneral   'DGeneral
'        lbCierreRealizado = oGen.GetCierreDiaRealizado(gdFecSis)
'        Set oGen = Nothing
'
'        If lbCierreRealizado Then
'            MsgBox "El cierre ya fue ralizado, no puede ingresar a esta opción.", vbInformation, "Aviso"
'            Exit Sub
'        End If
'        frmCredCargoAuto.Show 1
'    Case 2
'        FrmMantCargoAutomatico.Show 1
'End Select
'End Sub

'Private Sub M0301021600_Click(Index As Integer)
'Select Case Index
'    Case 1
'        frmCredAdmPrepago.Show 1
'    Case 2
'        frmCredMntPrepago.Show 1
'    Case 3 'JUEZ 20130925
'        frmCredExonerarPreCanc.Show 1
'End Select
'End Sub

'Private Sub M0301021700_Click(Index As Integer)
'    Select Case Index
'        Case 1
'            frmCredPolizas.Show 1
'        Case 2
'            frmCredPolizaGarantia.Inicio 1
'        Case 3
'            frmCredPolizasAprobacion.Show 1
'        Case 4
'            Call frmCredPolizaListado.Inicio(3)
'        Case 5
'            frmCredPoliTasas.Show 1
'        Case 6
'            frmRegistroPolizaCredVigente.Show 1
'        'WIOR 20140527 **************
'        Case 7
'            frmCredPolizaRenova.Show 1
'        'WIOR FIN *******************
'    End Select
'End Sub

'Private Sub M0301021800_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Pagos en Banco de la Nación
'            frmCredCorresponsaliaPagBcoNac.Show vbModal
'        Case 1 'Desembolso en Banco de la Nación
'            frmCredCorresponsaliaDesBcoNac.Show vbModal
'        Case 2
'            frmCredCorresponsaliaRecBcoNac.Show vbModal
'        'MADM 20101020
'        Case 3
'            frmCredPagoConvenioBNreporte.Show vbModal
'        'MADM ------
'    End Select
'End Sub

'Private Sub M0301021900_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmCredComite.Show 1
'        Case 1
'            frmCredComiteReporte.Show 1
'    End Select
'End Sub
'
'Private Sub M0301022000_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmSolicitudConstNoAdeudo.Show 1
'        Case 1
'            frmListarConstNoAdeudo.Show 1
'    End Select
'End Sub

'Private Sub M0301022100_Click(Index As Integer)
'    'JACA 20110628*********************************
'    Select Case Index
'        Case 0
'            'JUEZ 20121011 ************************
'            'frmCredRegistrarParametroBPPR.Show 1
'            'frmCredBPPTiposCartera.Show 1
'            'END JUEZ *****************************
'        Case 1
'            'frmCredAnalistaMetaBPP.Show 1
'        'JUEZ 20121011 ************************
'        Case 2
'            'frmCredBPPParametros.Show 1
'        'END JUEZ *****************************
'         'WIOR 20130702 *******************************
'        Case 3: frmCredBPPConfGen.Show 1
'        Case 4: frmCredBPPConfigMensual.Show 1
'         Case 6: frmCredBPPGeneracionesTotal.Show 1
'        'WIOR FIN ************************************
'    End Select
'    'JACA END****************************************
'End Sub

'WIOR 20130702 *******************************
'Private Sub M0301022106_Click(Index As Integer)
'    Select Case Index
'        Case 0: frmCredBPPMetasXAgencias.Show 1
'        Case 1: frmCredBPPConfigComiteCred.Show 1
'    End Select
'End Sub
'WIOR FIN ************************************

'WIOR 20140527 **************
'Private Sub M0301022108_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmCredBPPBonoPromotoresConf.Show 1
'        Case 1
'            frmCredBPPBonoPromotores.Show 1
'    End Select
'End Sub
'WIOR FIN *******************

'Private Sub M0301022200_Click(Index As Integer)
'    'BRGO 20111230*********************************
'    Select Case Index
'        Case 0
'
'        Case 1
'            frmCredPersonaCOFIDE.Inicio
'        Case 2
'            frmCredVehiculoInfoGas.Inicio
'        Case 3
'            frmINFOGASGeneracionXML.Show 1
'        Case 4
'            Call frmINFOGASLecturaArchivos.Inicio("05000", "Confirmación de habilitación de vehículo")
'        Case 5
'            Call frmINFOGASLecturaArchivos.Inicio("05001", "Recaudo de Vehículo")
'    End Select
'    'BRGO END****************************************
'End Sub
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
'Private Sub M0301022400_Click(Index As Integer)
'    Select Case Index
'        Case 1
'            frmCredSugAprob.Inicio 1, "Credito: Estado - Asesoria Legal"
'        Case 2
'            frmParametroRevisionExp.Show
'        Case 3
'            frmCredSugAprob.Inicio 2, "Credito: Estado - Supervición de Credito"
'    End Select
'End Sub
'END MIOL *******************************************************************

'WIOR 20130411 ***********************************
'Private Sub M0301022500_Click(Index As Integer)
'Select Case Index
'        Case 0: frmCredRiesgos.Show 1
'        Case 1: frmCredRiesgosInformeMod.Show 1
'        Case 2: frmConfirmaComunicacionCredRapiflash.Inicia 'FRHU 20140331 RQ14178
'End Select
'End Sub
'WIOR FIN ****************************************

'JUEZ 20140530 ***********************************
'Private Sub M0301022501_Click(Index As Integer)
'Select Case Index
'        Case 0: frmCredSolicLimSecEcon.Show 1
'        Case 1: frmCredSolicLimSecEconExt.Show 1
'End Select
'End Sub
'END JUEZ ****************************************

'WIOR 20130727*************************************
'Private Sub M0301022601_Click(Index As Integer)
'Select Case Index
'        Case 0: frmCredAgricoParam.Show 1
'    End Select
'End Sub
'WIOR FIN*************************************

'WIOR 20130924 ********************************
'Private Sub M0301022700_Click(Index As Integer)
'    Select Case Index
'         Case 0: frmCredRegVisitaAnalista.Show 1
'        Case 1: frmCredRegVisitaJefe.Show 1
'    End Select
'End Sub
'WIOR FIN *************************************
'RECO20140208 ERS002*******************************
'Private Sub M0301022800_Click(Index As Integer)
'      Select Case Index
'        Case 0
'            frmCredReasigInst.Show 1
'        Case 1
'            frmCredAsignarConvenio.Show 1
'        Case 2
'            frmCredRetiroConvenio.Show 1
'        Case 3 'RECO20141018
'            frmColHistoCambios.Show 1
'    End Select
'End Sub
'RECO FIN********************************************
''FRHU 20140324 ERS172-2013 RQ13875
'Private Sub M0301022900_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmProyeccionSeguimientoPorAgencia.Show 1
'        Case 1
'            frmProyectadoVsEjecutadoPorAgencia.Show 1
'    End Select
'End Sub
'FIN FRHU 20140324 RQ13875
'WIOR 20120905 **************************************************************
'Private Sub M0301024500_Click(Index As Integer)
'Select Case Index
'        Case 0
'            frmCredEvalSeleccion.Inicio 1
'        Case 1
'            frmCredEvalSeleccion.Inicio 2
'        Case 2
'            frmCredEvalExtornoVerif.Show 1
'    End Select
'End Sub
'WIOR FIN *******************************************************************

'Private Sub M0301030000_Click(Index As Integer)
'    Select Case Index
'        Case 1
'            frmColPRescateJoyas.Show 1
'        Case 3  'Adjudicacion
'            '*** PEAC 20080412
'            'frmColPAdjudicaLotes.Show 1
'        Case 5 ' Chafaloneo
'        Case 11 'RECO20150219
'            frmColPCampAdjudicados.Show 1
'    End Select
'End Sub

'Private Sub M0301030100_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Registro
'            frmColPRegContratoDet.Show 1
'        Case 1 'Mantenimiento
'            frmColPMantPrestamoPig.Show 1
'        Case 2 'Anulacion
'            frmColPAnularPrestamoPig.Show 1
'        Case 3 'Bloqueo
'            'frmColPBloqueo.Show 1
'        Case 4 'RECO20140208 ERS002**************************
'            frmColPRegContratoAmpliacion.Show 1
'    End Select
'End Sub

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

'Private Sub M0301030400_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmColPSubastaPrepara.Show 1
'        Case 1
'            frmColPSubastaProceso.Show 1
'    End Select
'End Sub
'
'Private Sub M0301030500_Click(Index As Integer)
'
'    Select Case Index
'        Case 0
'            'PEAC 20080605-descomentar para otra compilacion
'            'frmColPRetasacion.Show 1
'        Case 1
'            frmColPVentaLotePrepara.Show vbModal
'        Case 2
'            frmColPVentaLoteProceso.Show vbModal
'    End Select
'
'End Sub
'
'Private Sub M0301030600_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmColPMovs.Show 1
'        Case 1
'            frmColPContratosxCliente.Show 1
'        Case 3
'             frmColPRepo.Inicio 1
'        Case 4
'             frmColPRepo.Inicio 2
'        Case 5
'             frmColPRepo.Inicio 3
'
'    End Select
'End Sub
'
'Private Sub M0301040000_Click(Index As Integer)
'    Select Case Index
'    Case 4
'        'frmPigSeleccionVentaFundicion.Show 1
'    Case 5
'        frmPigConsulta.Inicio
'    Case 6
'        'frmPigFundicionJoya.Show 1
'    Case 7
'        'FrmPigRepValores.Show 1 ' X Memoria
'    Case 8
'        'frmPigEvaluacionMensualClientes.Show 1
'    End Select
'End Sub

'Private Sub M0301040100_Click(Index As Integer)
'    Select Case Index
'    Case 0
'        frmPigTarifario.Show 1
'    Case 1
'        'FrmPigClasificaCli.Show 1 'X Memoria
'    End Select
'End Sub

'Private Sub M0301040200_Click(Index As Integer)
'    Select Case Index
'    Case 0
'        frmPigRegContrato.Show 1
'    Case 1
'        'frmPigMantContrato.Show 1
'    Case 2
'        'frmPigAnularContrato.Show 1 'X Memor
'    Case 3
'        'FrmPigBloqueo.Show 1 'X Memor
'    End Select
'End Sub
'
'Private Sub M0301040300_Click(Index As Integer)
'    Select Case Index
'    Case 0
'        'frmPigProyeccionGuia.Show 1
'    Case 1
'        'frmPigDespachoGuia.Show 1 'X Memom
'    Case 2
'        'frmPigRecepcionValija.Show 1
'    End Select
'End Sub

'Private Sub M0301040400_Click(Index As Integer)
'
'    Select Case Index
'    Case 0
'        'frmPigRegistroRemate.Show 1
'    Case 1
'        'frmPigProcesoRemate.Show 1
'    Case 2
''        Dim oPigRemate As DPigContrato
''        Dim rs As Recordset
''
''        Set oPigRemate = New DPigContrato
''        Set rs = oPigRemate.dObtieneDatosRemate(oPigRemate.dObtieneMaxRemate() - 1)
''        If Not (rs.EOF And rs.BOF) Then
''            If CStr(rs!cUbicacion) = Right(gsCodAge, 2) Then
''                FrmPigVentaRemate.Show 1
''            Else
''                MsgBox "Usuario no se encuentra asignado en la Agencia de Remate", vbInformation, "Aviso"
''                Exit Sub
''            End If
''        End If
'    End Select
'End Sub
'
'Private Sub M0301030700_Click(Index As Integer)
'    Select Case Index
'        Case 0 'prepara adjudicacion
'            'PEAC 20080605-para desproteger en otra compilacion
'            frmColPAdjudicaPrepara.Show 1
'        Case 1 'adjudicacion de lotes
'            frmColPAdjudicaLotes.Show 1
'    End Select
'End Sub

'Private Sub M0301030800_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmPigMantenimientoPrecioTasacion.Show 1
'        Case 1 'RECO20140721 ERS114*************
'            frmColPTarifarioCartaNotarialMinka.Inicio 1, "Mantenimineto"
'    End Select
'End Sub
'
''RECO20140805 ERS074-2014*************************
'Private Sub M0301030900_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmColPPreparacionRetasacionVigDif.Show 1
'        Case 1
'            frmColPRetasacionConsulta.Show 1
'        Case 2
'            frmColPRetasacionVigenteDiferida.Show 1
'    End Select
'End Sub
''RECO FIN*****************************************

'Private Sub M0301050000_Click(Index As Integer)
'    Select Case Index
'        'Case 0 ' Ingreso a Recup de Otras Entidades
'        '    frmColRecIngresoOtrasEnt.Show 1
'        Case 2 ' Gastos en Recuperaciones
'            frmColRecGastosRecuperaciones.Show 1
'        Case 3 ' Metodo de Liquidacion
'            frmColRecMetodoLiquid.Show 1
'        Case 5 ' Pago
'            frmColRecPagoCredRecup.Inicio gColRecOpePagJudSDEfe, "PAGO CREDITO EN RECUPERACIONES", gsCodCMAC, gsNomCmac, True
'         '** Juez 20120418 **************************
'        Case 6 ' Cancelacion
'            'frmColRecCancelacion.Show 1
'            frmColRecCancPagoJudicial.Show 1
'        '** End Juez *******************************
'        Case 7 ' Castigo
'            frmColRecCastigar.Show 1
'        Case 8
'
'        Case 10
'            frmGarLevant.Show 1
'        Case 11
'            frmGarantExtorno.Show 1
'        Case 13
'            FrmColRecVistoRecup.Show 1
'        Case 14
'            'frmCredTransfRecupeGarant.Show 1
'        Case 15
'            'frmCredTransfGarantiaAdjudiSaneado.Show 1
'        Case 16
'            frmColBienesAdjudLista.Show 1
'        Case 17
'            'frmColEmbargadosListar.Show 1
'        'MADM 20111010
'        Case 18
'            FrmBloqueaRecupera.Show 1 'X Mem
'        Case 19 '*** PEAC 20120816
'            FrmColRecRegVisitaCliente.Show 1
'
'    End Select
'End Sub
'
'Private Sub M0301050100_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmColRecExped.Show 1
'        Case 1
'            frmColRecActuacionesProc.Inicia "N"
'    End Select
'End Sub

'Private Sub M0301050200_Click(Index As Integer)
' Dim oCredRel As New UCredRelac_Cli 'COMDCredito.UCOMCredRela   'UCredRelacion
' Select Case Index
'    Case 0
'        frmCredRelaCta.Inicio oCredRel, InicioMantenimiento, InicioRegistroForm
'        Set oCredRel = Nothing
'    Case 1
'        frmCredRelaCta.Inicio oCredRel, InicioMantenimiento, InicioConsultaForm
'        Set oCredRel = Nothing
'    Case 2
'        frmColRecComision.Show 1
'End Select
'End Sub

'Private Sub M0301050300_Click(Index As Integer)
'Select Case Index
'    Case 0
'        frmColRecRConsulta.Inicia "Consulta de Pagos de Créditos Judiciales"
'    Case 1
'        FrmColRecPagGestor.Show vbModal
'End Select
'End Sub
'
'Private Sub M0301050400_Click(Index As Integer)
'Select Case Index
'    Case 0
'        frmColRecReporte.Inicia "Reportes de Recuperaciones"
'End Select
'End Sub

'Private Sub M0301050500_Click(Index As Integer)
'Select Case Index
'    Case 0 ' Simulador
'        Call frmColRecNegCalculaCalendario.Inicio(1)
'    Case 1 'Registrar Negociacion
'        frmColRecNegRegistro.Inicia (True)
'    Case 2 'Resolver Negociacion
'        frmColRecNegRegistro.Inicia (False)
'End Select
'
'End Sub

'FRHU 20150428 ERS022-2015
'Private Sub M0301050600_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmColTransfCancelPago.Show 1
'    End Select
'End Sub
'FIN FRHU 20150428

'WIOR 20150602 ***
'Private Sub M0301052100_Click(Index As Integer)
'Select Case Index
'    Case 0
'        frmRecupCampConfig.Show 1
'    Case 1
'        frmRecupCampAcoger.Show 1
'    Case 2
'        frmRecupCampAuto.Show 1
'End Select
'End Sub
'WIOR FIN ********

'Private Sub M0301060000_Click(Index As Integer)
'Select Case Index
'    Case 1
'        FrmColocCalCargaRCC.Show 1
'    Case 3
'        FrmColocEvalRep.Show , Me
'    Case 4
'        FrmColocCalConsultaCliente.Show 1
'End Select
'End Sub
'
'Private Sub M0301060100_Click(Index As Integer)
'Select Case Index
'    Case 0
'        frmColocCalEvalCli.Inicio True
'    Case 1
'        frmColocCalEvalCliAutomatico.Show 1
'    Case 2 ' Este  formulario lo hizo Luis
'        FrmColocCalGarantiasPreferidas.Show 1
'    Case 3
'        FrmColocCalReclasificados.Show 1
'End Select
'End Sub
'
'Private Sub M0301060200_Click(Index As Integer)
'Select Case Index
'    Case 0
'        'ARCV 04-06-2007
'        'frmColocCalSist.Show 1
'        frmColocCalSist_NEW.Show 1
'    Case 1
'        frmAudCierraCalificacion.Show 1
'    Case 2
'        frmColocCalTabla.Show 1
'End Select
'
'End Sub

'Private Sub M0301060300_Click(Index As Integer)
'    Select Case Index
'        Case 1
'            'FrmIntDevTpoCred.Show 1
'        Case 2
'            'frmDistrExposicion.Show 1
'    End Select
'End Sub
'
'Private Sub M0301070000_Click(Index As Integer)
'Select Case Index
'    Case 0
'        frmRCDParametro.Show 1
'    Case 3
'        'frmRCDReporte.Show , Me
'        frmRCDReporte_NEW.Show , Me
'        'frmColocCalEvalCli.Inicio True
'    'ALPA 20120419***************
'    Case 5
'        frmRCDVcFecha.Show 1
'    '****************************
'    'ALPA 20120424***************
'    Case 6
'        frmRCDCambiarFormatoObs.Show 1
'    '****************************
'End Select
'
'End Sub
'
'Private Sub M0301070100_Click(Index As Integer)
'Select Case Index
'    Case 0 ' Datos Maestro RCD
'         frmRCDActualizaRCDMaestroPersona.Show 1
'    Case 1 ' Persona desde Maestro RCD
'        frmRCDActPersDeMaestroPersona.Show 1
'End Select
'End Sub
'
'Private Sub M0301070101_Click(Index As Integer)
'Select Case Index
'    Case 0 ' Carga de Errores
'        frmRCDErrorCargaTXT.Show 1
'    Case 1 ' Correccion de Errores
'        frmRCDErrorCorreccion.Show 1
'End Select
'End Sub

'Private Sub M0301070200_Click(Index As Integer)
'Select Case Index
'    Case 0 'Informe RCD
'        'frmRCDGeneraDatosRCD.Show 1 'ARCV 06-06-2007
'        frmRCDGeneraDatosRCD_NEW.Show 1
'    Case 1 ' Informe IBM
'        frmRCDGeneraDatosIBM.Show 1
'    Case 2
'        frmRCDVericaDatos.Show 1
'End Select
'End Sub
'
'Private Sub M0301070300_Click(Index As Integer)
'Select Case Index
'    Case 0 ' RCDMaestro Persona
'        frmRCDMantMaestroPersona.Show 1
'End Select
'End Sub
'
'Private Sub M0301080000_Click(Index As Integer)
'Select Case Index
'    Case 0
'        frmCredSolicitud.Inicioleasing (Registrar)
'    Case 1
'        frmCredSugerencia.InicioCargaDatos Registrar, True
'    Case 2
'        frmCredAprobacion.Aprobacion True
'    Case 3
'        frmCredDesembAbonoCta.DesembolsoCargoCuentaProveedorLeasing gCredDesembLeasing
'    Case 4
'        frmCredpagoCuotasLeasingDetalle.Inicia gCredPagLeasingCI
'    Case 5
'        frmUsuarioLeasing.Show 1
'End Select
'End Sub
''RECO20140208 ERS002***********************************************
'Private Sub M0301090000_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            If Not gnAgenciaHojaRutaNew Then
'                frmHojaRutaAnalista.Inicio 1, "- Registro"
'            Else
'                MsgBox "Ingrese a la opcion ''Colocaciones>Hoja Ruta>Resultado de Visita'' para registrar el resultado de tu visita (Agencia Configurada con la Nueva Hoja de ruta)."
'            End If
'        Case 1
'            If Not gnAgenciaHojaRutaNew Then
'                frmHojaRutaAnalista.Inicio 2, "- Mantenimiento" '& gsCodCargo
'            Else
'                MsgBox "Ingrese a la opcion ''Colocaciones>Hoja Ruta>Resultado de Visita'' para registrar el resultado de tu visita (Agencia Configurada con la Nueva Hoja de ruta)."
'            End If
'        Case 2
'            If Not gnAgenciaHojaRutaNew Then
'                frmHojaRutaAnalista.Inicio 3, "- Mantenimiento Coordinador" '& gsCodCargo
'            Else
'                MsgBox "Ingrese a la opcion ''Colocaciones>Hoja Ruta>Resultado de Visita'' para registrar el resultado de tu visita (Agencia Configurada con la Nueva Hoja de ruta)."
'            End If
'        Case 3
'            If Not gnAgenciaHojaRutaNew Then
'                frmHojaRutaAnalistaResultado.Inicio
'            Else
'                MsgBox "Ingrese a la opcion ''Colocaciones>Hoja Ruta>Resultado de Visita'' para registrar el resultado de tu visita (Agencia Configurada con la Nueva Hoja de ruta)."
'            End If
'        Case 4
'            If Not gnAgenciaHojaRutaNew Then
'                frmHojaRutaAnalistaConsulta.Show 1
'            Else
'                MsgBox "Ingrese a la opcion ''Colocaciones>Hoja Ruta>Resultado de Visita'' para registrar el resultado de tu visita (Agencia Configurada con la Nueva Hoja de ruta)."
'            End If
'        Case 5
'            If gnAgenciaHojaRutaNew Then
'                frmHojaRutaAnalistaGeneraResultado.Inicio
'            Else
'                MsgBox "Ingrese a las opciones anteriores para registrar el resultado de tu visita (Agencia No Configurada con la Nueva Hoja de ruta)."
'            End If
'        Case 6
'            If gnAgenciaHojaRutaNew Then
'                frmHojaRutaAnalistaDarVisto.Show 1
'            Else
'                MsgBox "Ingrese a las opciones anteriores para registrar el resultado de tu visita (Agencia No Configurada con la Nueva Hoja de ruta)."
'            End If
'    End Select
'End Sub
'RECO FIN***********************************************************
Private Sub M0401000000_Click(Index As Integer)
If Index = 2 Or Index = 3 Or Index = 4 Or Index = 9 Then
    Dim clsTC As COMDConstSistema.NCOMTipoCambio
    Dim nTC As Double
    Set clsTC = New COMDConstSistema.NCOMTipoCambio
    nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
    Set clsTC = Nothing
    If nTC = 0 Then
        MsgBox "NO se ha registrado el TIPO DE CAMBIO. Debe registrarse para iniciar operaciones.", vbInformation, "Aviso"
        Exit Sub
    End If
End If

'Dim sfiltro() As String
'Dim lnFiltraTC As Integer
'Dim lnFiltraMP As Integer
Dim oGen As COMDConstSistema.DCOMGeneral
Dim lbCierreRealizado As Boolean
Dim lbCierreCajaRealizado As Boolean
Dim oCaj As COMNCajaGeneral.NCOMCajero

Set oGen = New COMDConstSistema.DCOMGeneral
'lnFiltraTC = CInt(oGen.LeeConstSistema(102))
'lnFiltraMP = CInt(oGen.LeeConstSistema(103))
lbCierreRealizado = oGen.GetCierreDiaRealizado(gdFecSis)
Set oGen = Nothing

If Not lbCierreRealizado Then
    Set oCaj = New COMNCajaGeneral.NCOMCajero
    lbCierreCajaRealizado = oCaj.YaRealizoCierreAgencia(gsCodAge, gdFecSis)
    Set oCaj = Nothing
End If
    'RECO20151111 ERS061-2015******************
        If lbCierreCajaRealizado Then
            If Not VerificaGrupoPermisoPostCierre Then
                lbCierreCajaRealizado = True
            Else
                lbCierreCajaRealizado = False
            End If
        End If
    'RECO FIN *********************************
    Select Case Index
        Case 0
            frmMantTipoCambio.Show 1
        Case 2
            If lbCierreRealizado Then
                MsgBox "El cierre ya fue realizado, no puede ingresar a esta opción.", vbInformation, "Aviso"
                Exit Sub
            End If
            If lbCierreCajaRealizado Then
                MsgBox "El cierre de caja de la agencia ya fue ralizado, no puede ingresar a esta opción.", vbExclamation, "Aviso"
                Exit Sub
            End If
'            ReDim sfiltro(5)
'            If lnFiltraMP = 1 Then
'                sfiltro(1) = "[1][01234][012]" 'Pigno Trujillo
'            ElseIf lnFiltraMP = 2 Then
'                sfiltro(1) = "[1][01345][012]"   'Pigno Lima
'            ElseIf lnFiltraMP = 0 Then
'                sfiltro(1) = "[1][012345][012]"   'Ambos
'            End If
'
'            sfiltro(2) = "[23][0-2][0123]"    'Captaciones
'
'            If lnFiltraTC = 0 Then
'                sfiltro(3) = "90002[0-3]"       'Compra Venta
'            ElseIf lnFiltraTC = 1 Then
'                sfiltro(3) = "90002[0-6]"
'            End If
'            sfiltro(4) = "9010[01][0123456789]"    'Control de Efectivo Boveda y Cajero
'            sfiltro(5) = "90003[0-5]"    'Operaciones con Cheques
'            frmCajeroOperaciones.Inicia sfiltro, "Cajero - Operaciones"
            If gRsOpeF2.RecordCount = 0 Then
                MsgBox "El usuario no tiene permisos para esta opción", vbInformation, "Mensaje"
                Exit Sub
            End If
            'If Not VerificarRFIII Then Exit Sub ' *** RIRO SEGUN TI-ERS108-2013 ***
            gRsOpeF2.MoveFirst
            frmCajeroOperaciones.Inicia "Cajero - Operaciones", gRsOpeF2
            
        Case 3
            If lbCierreRealizado Then
                MsgBox "El cierre ya fue ralizado, no puede ingresar a esta opción.", vbInformation, "Aviso"
                Exit Sub
            End If
            If lbCierreCajaRealizado Then
                MsgBox "El cierre de caja de la agencia ya fue ralizado, no puede ingresar a esta opción.", vbExclamation, "Aviso"
                Exit Sub
            End If
'            ReDim sfiltro(4)
'            sfiltro(1) = "260[0-3]" 'Operaciones de Captaciones
'            sfiltro(2) = "126"      'Operaciones de Prendario
'            sfiltro(3) = "106"      'Operaciones de Colocaciones
'            sfiltro(4) = "136"      'Operaciones de judiciales
'            frmCajeroOpeCMAC.Inicia sfiltro, "Cajero - Operaciones CMACs Recepción"
            If gRsOpeCMACRecep.RecordCount = 0 Then
                MsgBox "El usuario no tiene permisos para esta opción", vbInformation, "Mensaje"
                Exit Sub
            End If
            gRsOpeCMACRecep.MoveFirst
            frmCajeroOpeCMAC.Inicia "Cajero - Operaciones CMACs Recepción", gRsOpeCMACRecep
        Case 4
            If lbCierreRealizado Then
                MsgBox "El cierre ya fue ralizado, no puede ingresar a esta opción.", vbInformation, "Aviso"
                Exit Sub
            End If
            If lbCierreCajaRealizado Then
                MsgBox "El cierre de caja de la agencia ya fue ralizado, no puede ingresar a esta opción.", vbExclamation, "Aviso"
                Exit Sub
            End If
'            ReDim sfiltro(3)
'            sfiltro(1) = "2605"     'Operaciones de Captaciones
'            sfiltro(2) = "127"      'Operaciones de Prendario
'            sfiltro(3) = "107"      'Operaciones de Colocaciones
'            frmCajeroOpeCMAC.Inicia sfiltro, "Cajero - Operaciones CMACs Llamada"
            If gRsOpeCMACLlam.RecordCount = 0 Then
                MsgBox "El usuario no tiene permisos para esta opción", vbInformation, "Mensaje"
                Exit Sub
            End If
            gRsOpeCMACLlam.MoveFirst
            frmCajeroOpeCMAC.Inicia "Cajero - Operaciones CMACs Llamada", gRsOpeCMACLlam
            frmCajeroOpeCMAC.Inicia "Cajero - Operaciones CMACs Recepción", gRsOpeCMACRecep
        Case 5
            If lbCierreRealizado Then
                MsgBox "El cierre ya fue ralizado, no puede ingresar a esta opción.", vbInformation, "Aviso"
                Exit Sub
            End If
            If lbCierreCajaRealizado Then
                MsgBox "El cierre de caja de la agencia ya fue ralizado, no puede ingresar a esta opción.", vbExclamation, "Aviso"
                Exit Sub
            End If
            If gRsOpeCMACLlam.RecordCount = 0 Then
                MsgBox "El usuario no tiene permisos para esta opción", vbInformation, "Mensaje"
                Exit Sub
            End If
            gRsOpeCMACLlam.MoveFirst
            'frmPITOperacionesInterCMAC.Inicia "Cajero - Operaciones InterCMACs (Envío)" ', gRsOpeInterCMACs
        Case 9
            If lbCierreRealizado Then
                MsgBox "El cierre ya fue realizado, no puede ingresar a esta opción.", vbInformation, "Aviso"
                Exit Sub
            End If
            If lbCierreCajaRealizado Then
                MsgBox "El cierre de caja de la agencia ya fue ralizado, no puede ingresar a esta opción.", vbExclamation, "Aviso"
                Exit Sub
            End If
'            ReDim sfiltro(9)
'            sfiltro(1) = "1[034][79][012]"         'Extornos de Colocaciones
'            If lnFiltraMP = 1 Then
'                sfiltro(2) = "129"          'Extornos de Prendario Trujillo
'            ElseIf lnFiltraMP = 2 Then      'Extornos de Prendario Lima
'                sfiltro(2) = "159"
'            ElseIf lnFiltraMP = 0 Then      'Extornos de Prendario Lima
'                sfiltro(2) = "1[25]9"
'            End If
'            sfiltro(3) = "2[3457]"      'Extornos de Captaciones
'            sfiltro(4) = "3[569]"       'Extornos de Otras Operaciones
'            If lnFiltraTC = 0 Then
'                sfiltro(5) = "90900[0-3]"
'            ElseIf lnFiltraTC = 1 Then
'               '  sfiltro(1) = "90900[0-6]"
'                sfiltro(5) = "90900[0-6]"
'            End If
'            sfiltro(6) = "90103[0-9]"   'Extornos de Operaciones de Boveda
'            sfiltro(7) = "90102[1-9]"   'Extornos de Operaciones de Cajero
'            sfiltro(8) = "90003[6-9]"   'Extornos de Operaciones con Cheque
'            sfiltro(9) = "90004[4-6]"   'Extornos de Compra Venta - Tipo de Cambio Especial
            
'            frmCajeroOperaciones.Inicia sfiltro, "Cajero - Extornos"
            If gRsExtornos.RecordCount = 0 Then
                MsgBox "El usuario no tiene permisos para esta opción", vbInformation, "Mensaje"
                Exit Sub
            End If
            'If Not VerificarRFIII Then Exit Sub ' *** RIRO SEGUN TI-ERS108-2013 ***
            gRsExtornos.MoveFirst
            frmCajeroOperaciones.Inicia "Cajero - Extornos", gRsExtornos
'        Case 12
'
'              frmExoneracionITF.Show
'
'           ' AvisoOperacionesPendientes
'
'        Case 13
'             frmCapNoCobroInactivas.Show 1
'
        '** Juez 20120723 *************
         Case 16
            frmOpeReimprVoucher.Show 1
        '** End Juez ******************
        
        Case 18
            frmActivacionPerfilRFIII.Show 1 ' *** RIRO SEGUN TI-ERS108-2013 ***
            
    End Select
End Sub

Private Sub M0401010000_Click(Index As Integer)
Select Case Index
    Case 0
        'Case gOpeHabCajDevABoveMN, gOpeHabCajDevABoveME
        frmCajeroHab.Show 1
        'Case gOpeHabCajTransfEfectCajerosMN, gOpeHabCajTransfEfectCajerosME
    Case 1
        frmCajeroHab.Show 1
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
       'Call frmCierreDiario.CierreDia 'frhu 2017
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
            frmCapExtornos.Show 1
        
            
        Case 1 '    Extorno Credito
        
        Case 2 '    Pignoraticio
            'frmColPExtornoOpe.Show 1
        Case 3 '    recuperaciones
            
    End Select
End Sub

Private Sub M0401040100_Click(Index As Integer)
    Select Case Index
        Case 0
            frmCredExtornos.Show 1
        Case 1
            frmCredExtornos.Show 1
        Case 2
            frmCredExtPagoLote.Show 1
    End Select
End Sub

Private Sub M0401050000_Click(Index As Integer)
    Select Case Index
    Case 0
        'frmAsientoDN.Inicio True
    Case 1
        'frmAsientoDN.Inicio False
    Case 2
        frmCtaContMantenimiento.Show 1
    End Select
End Sub

Private Sub M0401060000_Click(Index As Integer)
Dim sCad As String
Dim oPrevio As previo.clsprevio

Select Case Index
    Case 0
        gsOpeDesc = "RESUMEN DE INGRESOS Y EGRESOS CONSOLIDADO"
        frmCajeroIngEgre.Inicia False, True
    Case 1
        Dim orep As COMNCaptaGenerales.NCOMCaptaReportes
        
        Set orep = New COMNCaptaGenerales.NCOMCaptaReportes
        'madm 20101012 - parametro agencia
        sCad = orep.ReporteTrasTotSM("DETALLE DE OPERACIONES", False, gsCodUser, Format$(gdFecSis, "yyyymmdd"))
        Set orep = Nothing
        
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
        'COMENTADO POR PTI1 ACTA ACTA Nº 189-2018
        'Dim oProt As COMNCaptaGenerales.NCOMCaptaReportes
        'Set oProt = New COMNCaptaGenerales.NCOMCaptaReportes
        'sCad = oProt.ProtocoloOperaciones("PROTOCOLO DE USUARIO SOLES", 0, 0, gsNomAge, gcEmpresa, gdFecSis, gMonedaNacional, gsCodUser, , Format(gdFecSis, gsFormatoFechaView), gsCodAge)
        'sCad = sCad & oProt.ProtocoloOperaciones("PROTOCOLO DE USUARIO DOLARES", 0, 0, gsNomAge, gcEmpresa, gdFecSis, gMonedaExtranjera, gsCodUser, , Format(gdFecSis, gsFormatoFechaView), gsCodAge)
        
        'Set oPrevio = New previo.clsprevio
        'oPrevio.Show sCad, "PROTOCOLO DE USUARIO", True
        'Set oPrevio = Nothing
        'FIN COMENTADO PTI1
    Case 4
        frmOperacionesNum.Show 1
    Case 6
        FrmITFGeneraArchivos.Show 1
    Case 8
        sCad = RepHavDevBoveda(gdFecSis, gsNomCmac, gsNomAge, gsCodAge)
        If sCad = "" Then
        MsgBox "No existe información"
        End If
    Case 9 '**DAOR 20080125, Reporte de Registros de Efectivo por usuario
        frmOpeReportes.Show vbModal
        'MADM 20101019
    Case 10
        frmCajeroIngDetalleGral.Show vbModal
    Case 11 'JUEZ 20130601
        frmEnvioEstadoCtaRep.Show 1
    Case 12 'JUEZ 20131021
        frmOpeRepActDatosCamp.Show 1
End Select
End Sub

 Public Function RepHavDevBoveda(psFecSis As Date, psNomCmac As String, psNomAge As String, psCodAge As String) As String
  Dim cMovNro As String, rstemp As ADODB.Recordset
  Dim orep As nCaptaReportes
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
  
  
  Set orep = New nCaptaReportes
    Set rstemp = orep.RepHavDevBoveda(Format(psFecSis, "yyyymmdd"), psCodAge)
   
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
                xlHoja1.Cells(nFila, 2) = rstemp!nMoneda
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
                xlHoja1.Cells(nFila, 2) = rstemp!nMoneda
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
            
            Set rstemp = orep.REPBOVSALDOS(psCodAge, Format(psFecSis, "yyyymmdd"), Format(psFecSis, "yyyymmdd"))
            
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
                xlHoja1.Cells(nFila, 5) = Format(CDate(Mid(rstemp!dFecha, 5, 2) & "/" & Right(rstemp!dFecha, 2) & "/" & Left(rstemp!dFecha, 4)), "dd/MM/yyyy")
                               
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
  MsgBox err.Description, vbInformation, "Aviso"
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
   MsgBox err.Description, vbInformation, "Aviso"
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



Private Sub M0501020000_Click(Index As Integer)

End Sub

 



Private Sub M0401070000_Click(Index As Integer)
Select Case Index
    Case 1
        frmCajeroCierreAgencias.Show 1
    Case 2
        gsOpeCod = gOpeCajaExtCierreAgenica
        gsOpeDesc = "EXTORNO CIERRE CAJA AGENCIA"
        frmCajeroExtornos.Inicia "CAJERO - EXTORNO CIERRE CAJA AGENCIA"
End Select
End Sub

Private Sub M0401080000_Click(Index As Integer)
Select Case Index
    Case 0
        frmCajeroBilletajeAutomatico.Inicio 1
    Case 1
        frmCajeroBilletajeAutomatico.Inicio 2
End Select
End Sub

'** Juez 20120807 ******************************
Private Sub M0401090000_Click(Index As Integer)
    Dim VistoCorrecto As Boolean
    'RIRO20140902 ***********
    If Index = 0 Or Index = 1 Or Index = 3 Then
        Dim clsTC As COMDConstSistema.NCOMTipoCambio
        Dim nTC As Double
        Set clsTC = New COMDConstSistema.NCOMTipoCambio
        nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
        Set clsTC = Nothing
        If nTC = 0 Then
            MsgBox "NO se ha registrado el TIPO DE CAMBIO. Debe registrarse para iniciar operaciones.", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    'END RIRO ***************
    
    Select Case Index
        Case 0
            frmCajaArqueoVentBov.Inicia 1 'Ventanilla
        Case 1
            frmCajaArqueoVentBov.Inicia 2 'Boveda
        Case 2
            frmCajaArqueoVentBovExt.Show 1
        'RIRO20140630 ERS072
        Case 3
            frmCajaArqueoVentBov.Inicia 3 'Entre Ventanillas
        'END RIRO
        Case 4 'PASI20151219
            frmArqueoTarjDebBoveda.Inicia
        Case 5
            frmArqueoTarjDebVent.Inicia
        Case 6
            frmExtornoArqueoTarjDebBoveda.Inicia 0
        Case 7
            frmExtornoArqueoTarjDebBoveda.Inicia 1
        'end PASI
         'ANDE 20170529 ERS021-2017
        Case 8
            VistoCorrecto = frmVistoElectronico.Inicio(19)
            If VistoCorrecto Then
                frmArqueoExpedientesAho.Show 1
            End If
        Case 9
            VistoCorrecto = frmVistoElectronico.Inicio(20)
            If VistoCorrecto Then
                frmExtornoArqueoExpAho.Show 1
            End If
        'END ANDE
    End Select
End Sub
'** End Juez ***********************************

'FRHU 20140505 ERS063-2014
Private Sub M0401100000_Click(Index As Integer)
    Select Case Index
        Case 0 'Aprobacion / Rechazo Operaciones
            frmCapRegAproAutOtrasOperaciones.Show 1
        '***MARG 20171222-ers065-2017 ---SUBIDO DESDE LA 60***
        Case 1
            frmCapConfirmacionAtencionSinTarjeta.Show 1
        '***end MARG***********
        Case 2 'CROB20180531
            frmAprobarCreditoPigno.Show 1 'CROB20180531
        '**ARLO20180605****
        Case 3
            frmPreDesemyExtornoCompraDeuda.Inicio (1)
        Case 4
            frmPreDesemyExtornoCompraDeuda.Inicio (2)
        '**ARLO20180605****
    End Select
End Sub
'FIN FRHU 20140505

'RIRO 20150701 ERS162-2014
Private Sub M0401110000_Click(Index As Integer)
    Dim oUtilidades As New frmUtilidadesTrama
    Select Case Index
        Case 0
            oUtilidades.Inicia 1
        Case 1
            oUtilidades.Inicia 2
    End Select
End Sub
'END RIRO
'VAPA20170707 CCE
Private Sub M0401120000_Click(Index As Integer)
 Select Case Index
    Case 0
       frmCCEGeneracionArch.Show 1
    Case 1
        frmCCECargaArch.Show 1
    Case 2
        frmCCECargaOficinas.Show 1
    Case 3
        frmCCEReporteSaldos.Show 1
    Case 4
       Dim obj As frmCCEExtornoTrama
       Set obj = New frmCCEExtornoTrama
       obj.Show 1
       Set obj = Nothing
End Select
End Sub
'END VAPA
'APRI20180407 ERS036-2017
Private Sub M0401130000_Click(Index As Integer)
    Select Case Index
        Case 1:
            frmEnvioEstadoCtaAfiliacion.Inicia Index
        Case 2:
            frmEnvioEstadoCtaAfiliacion.Inicia Index
    End Select
End Sub
'END APRI

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
'        Case 5 'Ctas Contables
'            frmCtaContMantenimiento.Show 1
'        Case 6 'backUp
'            'frmBackUp.Show 1
'        Case 7
'            frmCajeroGrupoOpe.Show 1
'        'Case 8
'         '   frmCapMantOperacion.Show 1
'        'Case 9
'         '   frmMantCodigoPostal.Show 1
'        Case 10
'            frmDocRecParam.Show 1
'        'Case 11
'          '  frmMantCIIU.Show 1
'        Case 12
'            FrmMantFeriados.Show 1


        Case 3
            frmHabilitacionOpciones.Show 1

    End Select
End Sub
''***Agregado por ELRO el 20120720, según OYP-RFC077-2012
'Private Sub M0601010000_Click(Index As Integer)
'    'frmLimEfeRegistro.Show 1
'    frmLimiteEfectivoAdm.Inicio 1, "Registro" 'RECO20150206 ERS-022-2014
'End Sub
''***Fin Agregado por ELRO el 20120720*******************

'Private Sub M0601020000_Click(Index As Integer)
'    'frmLimEfeMantenimiento.Show 1
'    frmLimiteEfectivoAdm.Inicio 2, "Mantenimiento" 'RECO20150206 ERS-022-2014
'End Sub

'Private Sub M0701000000_Click(Index As Integer)
'    If Index = 3 Then
'        frmPosicionCli.Show 1
'    End If
'    'If Index = 6 Then  ' Reportes
'    '    FrmPersReporte.Show 1
'    'End If
'End Sub
'
'Private Sub M0701010000_Click(Index As Integer)
'    'Persona
'    Select Case Index
'        Case 0 'Registro
'            frmPersona.Registrar
'        Case 1 'mantenimiento
'            frmPersona.Mantenimeinto
'        Case 2 'Consulta
'            frmPersona.Consultar
'        Case 3 'Exoneradas del Lavado de Dinero
''            frmPersLavDinero.Show 1
'        Case 4 'Rol de Persona
''            FrmPersonaRolMantenimiento.Show 1
'        Case 5
''            frmPersComentario.Show 1
'        Case 6
'            frmPersGrupoE.Show 1
'        'By Capi 30012008
'        Case 7 'Dudosas del Lavado de Dinero
'            frmPersLavDineroDudoso.Show 1
'
'        '*** PEAC 20090715
'        Case 8 'Registro Clientes lista negativa
'            frmPersNegativas.Show 1
'        'MADM 20100524 - Autorizacion Lista Negativa
'        Case 9
'            'frmPersNegativaAutorizacion.Show 1'WIOR 20121123 COMENTO
'            frmPersNegativaAutorizacion.Inicio 'WIOR 20121123
'    'WIOR 20121123 PRE AUTORIZACION ***********************************
'       Case 10
'            frmPersNegativaAutorizacion.Inicio True
'    'WIOR FIN *********************************************************
'        Case 11 'WIOR 20121123 cambio de 10 a 11
'            frmPersAdministrarSesiones.Show 1
'        Case 12 'WIOR 20121123 cambio de 11 a 12
'            frmCapAbonoPersParam.Show 1 'JACA  20110320
'        Case 13 'WIOR 20121123 cambio de 12 a 13
'            frmCredPagoCuotasPersParam.Show 1 '** Juez 20120514
'        Case 14 '** Modif Juez 20120514 Se cambio nro 12 >> 13'WIOR 20121123 cambio de 13 a 14
'            frmPersClienteSensible.Show  'Modificado ELRO
'        Case 15 'JUEZ 20130717
''            frmPersPREDA.inicio
'        'WIOR 20140107 ******************
'        Case 17
'            frmCredHonrados.Show 1
'        'WIOR FIN ***********************
'        'FRHU 20140401 RQ14132 **********
'        Case 18
'            frmPersBusqueda.Show 1
'        'FIN FRHU 20140401 **************
'        Case 19 'FRHU 20150310 ERS013-2015
'            frmPersEstadosFinancieros.Show 1
'     End Select
'End Sub

'JUEZ 20131016 *******************************************
'Private Sub M0701010100_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmPersonaCampDatosBitacora.InicioActualizar
'        Case 1
'            frmPersonaCampDatosBitacora.InicioConsultar
'    End Select
'End Sub
'END JUEZ ************************************************

'Private Sub M0701020000_Click(Index As Integer)
'    'Instituciones Financieras
'    Select Case Index
'        Case 0
'            frmMntInstFinanc.InicioActualizar
'        Case 1
'            frmMntInstFinanc.InicioConsulta
'    End Select
'End Sub

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

'Private Sub M100102010000_Click(Index As Integer)
'    frmRevisionRegistrar.Show
'End Sub

'Private Sub M100102020000_Click(Index As Integer)
'    frmListarRevision.Show
'End Sub

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

'Private Sub M100201000000_Click(Index As Integer)
'    frmGenerarCarta.Show 1
'End Sub

Private Sub M100202000000_Click(Index As Integer)
'    frmCapReportes.Show
'    frmCapReportes.Inicializar_OperacionAhorros
End Sub

Private Sub M100203000000_Click(Index As Integer)
'    frmCapReportes.Show
'    frmCapReportes.Inicializar_CtaAperturadas
End Sub

Private Sub M100204000000_Click(Index As Integer)
    gRsOpeF2.MoveFirst
    frmCajeroOperaciones.Inicia "Cajero - Operaciones", gRsOpeF2
    gRsOpeF2.MoveFirst
    frmCajeroOperaciones.Inicializar_MovimientosCtasAhorros "Cajero - Operaciones", gRsOpeF2
End Sub

Private Sub M1001000000_Click(Index As Integer)
    frmAdmCredRegVisitas.Show 1
End Sub

'RECO20150318 ERS010-2015************************
Private Sub M1002010000_Click(Index As Integer)
    Select Case Index
        Case 0
            frmAdmControlCredDesemb.Show 1
        Case 1
            frmAdmCredRegControl.Inicio "Post-Desembolso", "", 2
        Case 2
            'frmAdmConfigCheckList.Show 1 'Descomentado by NAGL Según RFC1712290003 ' Comentado por Catálogo Productos yihu20190215
    'ACTA Nº 113 - 2019 - JOEP - SORO - Asignacion de Agencias para Pre-Desembolsos
        Case 3
           frmAdmAsignarUsuarioPreDes.Show 1
    'ACTA Nº 113 - 2019 - JOEP - SORO - Asignacion de Agencias para Pre-Desembolsos
    End Select
End Sub
'RECO FIN***************************************
'Private Sub M100301000000_Click(Index As Integer)
'    'frmGenerarCartaCredito.Show 1
'End Sub
'
'Private Sub M100302000000_Click(Index As Integer)
''    frmCredReportes.inicia "Reportes de Créditos"
''    frmCredReportes.Inicializar_OperacionesReprogramadas ("108325")
'End Sub

Private Sub M1003000000_Click(Index As Integer)
    frmReportesAdmCred.Show 1
End Sub

Private Sub M1004000000_Click(Index As Integer)
    frmAdmCredAutoMant.Show 1
End Sub

Private Sub M1005000000_Click(Index As Integer)
    frmAdmCredExoMant.Show 1
End Sub
'WIOR 20120616 ************************************************
Private Sub M1006020000_Click(Index As Integer)
    Select Case Index
        'Hojas CF
        Case 0 'REMESAR
            frmCFHojasRemesar.Show 1
        Case 1 'RECEPCIONAR
            frmCFHojasRecepcion.Show 1
        Case 2 'CONSULTAR
            frmCFHojasConsultar.Show 1
    End Select
End Sub
'WIOR FIN ***************************************************
'WIOR 20140128 ************************************************
Private Sub M1007030000_Click(Index As Integer)
    Select Case Index
        'Créditos Vinculados
        Case 0 'Asignacion Saldos[Créditos]
            frmCredSaldosVincAsignar.Inicio (1)
        Case 1 'Asignacion Saldos[Ventanilla]
            frmCredSaldosVincAsignar.Inicio (2)
        Case 2 'Estado Actual
            'frmCredSaldosVincEstado.Show 1
            frmCredSaldosVincEstado.Inicio 'ORCR20140314
        Case 3 'Saldo Disponible
            'frmCredSaldosVincColaborador.Show 1
            frmCredSaldosVincColaborador.Inicio 'ORCR20140314
        Case 4 'Patrimonio Efectivo Ajustado
            frmPatrimonioEfectivo.Show 1
    End Select
End Sub
'WIOR FIN ***************************************************
'RECO20150326 ERS010-2015************************************
Private Sub M1008040000_Click(Index As Integer)
    Select Case Index
        Case 0
            frmAdmCredAutorizacionChkList.Show 1
        Case 1
            'frmAdmCredChekListMant.Inicio "Mantenimiento", 1 'Quitado por CatálogoProductos yihu20190215
        Case 2
            'frmAdmCredChekListMant.Inicio "Consulta", 2
    End Select
End Sub
'marg--
Private Sub M1009050000_Click(Index As Integer)
    frmArqueoPagare.Inicia
End Sub
'end marg--
'PASI20171216 x Migracion***
Private Sub M1103000000_Click(Index As Integer)
    frmDJSujetosObligados.Show 1
End Sub
'PASI END***
'Private Sub M100303010000_Click(Index As Integer)
'    frmCredMntGastos.Inicio InicioGastosConsultar
'End Sub
'PASI20171215 x Migración***
Private Sub M1104010000_Click(Index As Integer)
    frmOCNivelesRiesgoConfig.Show 1
End Sub
'PASI END***
'Private Sub M100303020000_Click(Index As Integer)
''    frmAuditReporteTarifario.Show 1
'End Sub
'PASI20171215 x Migración***
Private Sub M1105000000_Click(Index As Integer)
    frmOCNivelesRiesgoPorActividad.Show 1
End Sub
'PASI END***
'PASI20180604 x Migración ***
Private Sub M1201010000_Click(Index As Integer)
    Select Case Index
        Case 1
            'frmSegTarjetaSolicitud.Show 1 'FRHU 20140610 ERS068-2014
        Case 2
            frmSegTarjetaAfiliacionAnulacion.Show 1
        Case 3
            'frmSegTarjetaRechazoSolicitud.Show 1 'FRHU 20140610 ERS068-2014
        Case 4
            'frmSegTarjetaAceptacionSolicitud.Show 1 'FRHU 20140610 ERS068-2014
        Case 5 'JUEZ 20140615
            'frmSegTarjetaGeneraTramas.Show 1
        Case 6 'JUEZ 20150510
            frmSegTarjetaAnulaDevPend.Show 1
    End Select
End Sub
'PASI END***
'Private Sub M100304000000_Click(Index As Integer)
'    'frmAuditReporteCreditosDC.Show 1
'End Sub
'
'Private Sub M100401000000_Click(Index As Integer)
'    'frmBalanceHisto.Inicio 1
'End Sub
'
'Private Sub M100402000000_Click(Index As Integer)
''    gbBitCentral = True
''    frmReportes.Inicio "76", 1
''    frmReportes.Inicializar_operacion
'End Sub
'
'Private Sub M100501010000_Click(Index As Integer)
'    'gbBitCentral = True
'    'frmLogOCAtencion.Inicio True, "501210", False, False, True
'End Sub
'
'Private Sub M100501020000_Click(Index As Integer)
'    'gbBitCentral = True
'    'frmLogOCAtencion.Inicio True, "502210", False, False, True
'End Sub
'
'Private Sub M100501030000_Click(Index As Integer)
'    'gbBitCentral = True
'    'frmLogOCAtencion.Inicio False, "501211", False, False, True
'End Sub
'
'Private Sub M100501040000_Click(Index As Integer)
'    'gbBitCentral = True
'    'frmLogOCAtencion.Inicio False, "502211", False, False, True
'End Sub
'
'Private Sub M100502010000_Click(Index As Integer)
'    'gbBitCentral = True
'    'frmLogOCAtencion.Inicio True, "501210", False, False, True, True
'End Sub
'
'Private Sub M100502020000_Click(Index As Integer)
'    'gbBitCentral = True
'    'frmLogOCAtencion.Inicio True, "502210", False, False, True, True
'End Sub
'
'Private Sub M100502030000_Click(Index As Integer)
'    'gbBitCentral = True
'    'frmLogOCAtencion.Inicio False, "501211", False, False, True, True
'End Sub
'
'Private Sub M100502040000_Click(Index As Integer)
'    'gbBitCentral = True
'    'frmLogOCAtencion.Inicio False, "502211", False, False, True, True
'End Sub
'
'Private Sub M100601000000_Click(Index As Integer)
''    gbBitCentral = True
''    frmReportes.Inicio "76", 2
''    frmReportes.Inicializar_Operacion_PagoProveedores
'End Sub
'
'Private Sub M100602000000_Click(Index As Integer)
''    gbBitCentral = True
''    frmReportes.Inicio "4[6-9]", 3
''    frmReportes.Inicializar_Operacion_ConsultaSaldos
'End Sub
'
'Private Sub M100603000000_Click(Index As Integer)
''    gbBitCentral = True
''    frmReportes.Inicio "4[6-9]", 4
''    frmReportes.Inicializar_Operacion_ConsultaSaldosAdeudos
'End Sub
'
'Private Sub M100701000000_Click(Index As Integer)
'    'frmAuditListarAcceso.Show 1
'End Sub
'
'Private Sub M100801000000_Click(Index As Integer)
'    frmAudRegistroActividadProgramada.Show 1
'End Sub
'Private Sub M100802000000_Click(Index As Integer)
'    frmAudRegistroProcedimiento.Show 1
'End Sub
'PASI20180604 x Migración***
Private Sub M1201010100_Click(Index As Integer)
'   Select Case index
'        Case 0
'            frmSegTarjetaConfigParametros.Show 1
'        Case 1
'            frmSegTarjetaNCertificadosxAgencia.Show 1
'        Case 2
'            frmSegTarjetaConfigDoc.Show 1 'FRHU 20140610 ERS068-2014
'    End Select
End Sub
'end PASI***
'Private Sub M100803000000_Click(Index As Integer)
'    frmAudDesarrolloProcedimiento.Show 1
'End Sub
'
'Private Sub M100804000000_Click(Index As Integer)
'    frmAudDesarrolloProcedimientoVerificar.Show 1
'End Sub
'
'Private Sub M100901000000_Click(Index As Integer)
''    frmAudReporteSeguimientoActividades.Show 1
'End Sub
'
'Private Sub M110100000000_Click(Index As Integer)
'    frmAdmCredRegVisitas.Show 1
'End Sub
'
'Private Sub M110200000000_Click(Index As Integer)
'    'frmAdmCredRegControlCreditos.Show 1 'RECO20150421 ERS010-2015
'End Sub
'
''RECO20150318 ERS010-2015************************
'Private Sub M110201000000_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmAdmControlCredDesemb.Show 1
'        Case 1
'            frmAdmCredRegControl.Inicio "Post-Desembolso", "", 2
'        'Case 2
'           ' frmAdmConfigCheckList.Show 1
'    End Select
'End Sub
''RECO FIN***************************************
'
'Private Sub M110300000000_Click(Index As Integer)
'   'frmReportesAdmCred.Show 1
'End Sub
'
'Private Sub M110400000000_Click(Index As Integer)
'frmAdmCredAutoMant.Show 1
'End Sub
'
'Private Sub M110500000000_Click(Index As Integer)
'frmAdmCredExoMant.Show 1
'End Sub
''WIOR 20120616 ************************************************
'Private Sub M110601000000_Click(Index As Integer)
'    Select Case Index
'        'Hojas CF
'        Case 0 'REMESAR
'            frmCFHojasRemesar.Show 1
'        Case 1 'RECEPCIONAR
'            frmCFHojasRecepcion.Show 1
'        Case 2 'CONSULTAR
'            frmCFHojasConsultar.Show 1
'    End Select
'End Sub
''WIOR FIN ***************************************************
'PASI20180604 x Migración
Private Sub M1202010100_Click(Index As Integer)
'    Select Case index
'        Case 0
'            frmSegSolicitudCobertura.Inicia 1, gContraIncendio 'RECO20160214 ERS073-2015
'        Case 1
'            frmSegSolicitudCobertura.Inicia 2, gContraIncendio 'RECO20160214 ERS073-2015
'        Case 2
'            frmSegSolicitudCobertura.Inicia 3, gContraIncendio 'RECO20160214 ERS073-2015
'    End Select
End Sub
'end PASI***
'WIOR 20140128 ************************************************
'Private Sub M110701000000_Click(Index As Integer)
'  Select Case Index
'        'Créditos Vinculados
'        Case 0 'Asignacion Saldos[Créditos]
'            frmCredSaldosVincAsignar.Inicio (1)
'        Case 1 'Asignacion Saldos[Ventanilla]
'            frmCredSaldosVincAsignar.Inicio (2)
'        Case 2 'Estado Actual
'            'frmCredSaldosVincEstado.Show 1
'            frmCredSaldosVincEstado.Inicio 'ORCR20140314
'        Case 3 'Saldo Disponible
'            'frmCredSaldosVincColaborador.Show 1
'            frmCredSaldosVincColaborador.Inicio 'ORCR20140314
'        Case 4 'Patrimonio Efectivo Ajustado
'            frmPatrimonioEfectivo.Show 1
'    End Select
'End Sub
'WIOR FIN ***************************************************
'RECO20150326 ERS010-2015************************************
'Private Sub M110801000000_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmAdmCredAutorizacionChkList.Show 1
'        Case 1
'            frmAdmCredChekListMant.Inicio "Mantenimiento", 1
'        Case 2
'            frmAdmCredChekListMant.Inicio "Consulta", 2
'    End Select
'End Sub
'RECO FIN ****************************************************
'JACA 20110514***********************************************
'Private Sub M120100000000_Click(Index As Integer)
'    'frmPersOpeAgeOcupacion.Show 1
'End Sub
''JACA END****************************************************
'
'Private Sub M120200000000_Click(Index As Integer)
'    'JACA 20110530
'    frmOpeInusuales.Show 1
'End Sub
'
'Private Sub M120300000000_Click(Index As Integer)
'    frmDJSujetosObligados.Show 1
'End Sub
'FRHU 20140917 ERS106-2014
'Private Sub M120401000000_Click(Index As Integer)
'    frmOCNivelesRiesgoConfig.Show 1
'End Sub
'Private Sub M120500000000_Click(Index As Integer)
'    frmOCNivelesRiesgoPorActividad.Show 1
'End Sub
'FIN FRHU 20140917
'Private Sub M130100000000_Click(Index As Integer)
'    frmGiroMantenimiento.Show 1 '*** PEAC 20130222
'End Sub
'PASI20180604 x Migración
Private Sub M1202020100_Click(Index As Integer)
'    Select Case index
'        Case 0
'            frmSegRechazoSolicitud.Show 1
'        Case 1
'            frmSegResolucionSolicCober.Show 1
'        Case 2
'            frmSegSolicReconsideracion.Show 1
'        Case 3
'            frmSegExtornoSolicCober.Show 1
'    End Select
End Sub
'End PASI***
'Private Sub M130200000000_Click(Index As Integer)
'    frmMantCreditos.Show 1 '*** PEAC 20130222
'End Sub
'FRHU 20140528 ERS068-2014
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
'Private Sub M140101000000_Click(Index As Integer)
'    Select Case Index
'        Case 1
'            'frmSegTarjetaSolicitud.Show 1 'FRHU 20140610 ERS068-2014
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
'PASI20180604 x Migración***
Private Sub M1202030000_Click(Index As Integer)
    'frmSegHistorialDocumentos.Show 1
End Sub
'end PASI***
'FIN FRHU 20140528 ERS068-2014
'****>MAVM
'RECO20150326 ERS149-2014*************************
'Private Sub M140102000000_Click(Index As Integer)
'    Select Case Index
'        Case 2
'            frmSegHistorialDocumentos.Show 1
'    End Select
'End Sub
'Private Sub M140102010000_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmSegSolicitudCobertura.inicia 1
'        Case 1
'            frmSegSolicitudCobertura.inicia 2
'        Case 2
'            frmSegSolicitudCobertura.inicia 3
'    End Select
'End Sub
'PASI20180604 x Migración***
Private Sub M1203010000_Click(Index As Integer)
'    Select Case index
'        Case 0
'            frmSegParamSolicitud.Show 1
'    End Select
End Sub
'end PASI***
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
'
'Private Sub M140103000000_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmSegParamSolicitud.Show 1
'    End Select
'End Sub
'RECO FIN*****************************************
'Private Sub M1000000000_Click(Index As Integer)
'    FrmCredTraspCartera.Show 1
'End Sub
'
'Private Sub M140201000000_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmSegSepelioDesactivar.Show 1
'    End Select
'End Sub
'
'Private Sub M140201001000_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmSegParamSepelio.Show 1
'        Case 1
'            frmSegSepelioAfiliacion.IniciaAfilManual
'    End Select
'End Sub
'PASI20180604 x Migración ***
Private Sub M1204010000_Click(Index As Integer)
    Select Case Index
        Case 0
            frmSegSepelioDesactivar.Show 1
        Case 2
            'frmSegSolicitudCoberturaSepelio.Inicia 1
        Case 3
            'frmSegRechazoSolicitud.Inicia gSepelio
        Case 4
            'frmSegSepelioAceptacion.Show 1
        Case 5 'RECO20160425
            'frmSegSepelioActDatos.Show 1
    End Select
End Sub
'end PASI***
'PASI20180604 x Migración ***
Private Sub M1204010100_Click(Index As Integer)
'    Select Case index
'        Case 0
'            frmSegParamSepelio.Show 1
'    End Select
End Sub
'end PASI***
'PASI20180604 x Migración ***
Private Sub M1205010000_Click(Index As Integer)
'    Select Case index
'        Case 0
'            'frmSegGeneracionTrama.Show
'        Case 3
'            frmSegHistorialDocumentos.Inicia gMYPE
'    End Select
End Sub
'end PASI***
'PASI20180604 x Migración***
Private Sub M1205010100_Click(Index As Integer)
'    Select Case index
'        Case 0
'            frmSegSolicitudCobertura.Inicia 1, gMYPE
'        Case 1
'            frmSegSolicitudCobertura.Inicia 2, gMYPE
'        Case 2
'            frmSegSolicitudCobertura.Inicia 3, gMYPE
'    End Select
End Sub
'end PASI***
'PASI20180604 x Migración ***
Private Sub M1205010200_Click(Index As Integer)
'    Select Case index
'        Case 0
'            frmSegRechazoSolicitud.Inicia gMYPE
'        Case 1
'            frmSegResolucionSolicCober.Inicia gMYPE
'        Case 2
'            frmSegSolicReconsideracion.Inicia gMYPE
'        Case 3
'            frmSegExtornoSolicCober.Inicia gMYPE
'    End Select
End Sub
'end PASI***
'PASI20180604 x Migración***
Private Sub M1206000000_Click(Index As Integer)
    'frmGestionSiniestro.Show 1
End Sub
'end PASI***
Private Sub MDIForm_Click()
'  Form1.Show
End Sub

Private Sub MDIForm_Load()
 Me.Icon = LoadPicture(App.Path & gsRutaIcono)
Timer1.Enabled = False
CargaMensajes  'WIOR 20130826

'->***** LUCV20190323, Según RO-1000373
'Quita el borde de los dos controles
Image1.BorderStyle = 0
pbxFondo.BorderStyle = 0
'->***** Fin LUCV20190323

'RECO20140124 ERS154********************************************************
If gsCodCargo = "005002" Or gsCodCargo = "005003" Or gsCodCargo = "005004" Or gsCodCargo = "005005" Then

    If Not 0 Then 'gnAgenciaHojaRutaNew Then
    'Comentado por VAPI segun ERS0232015
        Dim oCred As New COMDCredito.DCOMCreditos
        Dim oDrNumEnDia As New ADODB.Recordset
        Dim oDrNumDiaAnt As New ADODB.Recordset
    
        Set oCred = New COMDCredito.DCOMCreditos
        Set oDrNumEnDia = New ADODB.Recordset
        Set oDrNumDiaAnt = New ADODB.Recordset
    
        Set oDrNumEnDia = oCred.ObtieneNumeroRutaEnDia(Format(gdFecSis, "yyyy/MM/dd"), gsCodUser)
        Set oDrNumDiaAnt = oCred.ObtieneNumeroRutaDiaAnt(Format(gdFecSis, "yyyy/MM/dd"), gsCodUser)
        'Set oDrNumEnDia = oCred.ObtieneNumeroRutaEnDia(Format(gdFecSis, "yyyy/MM/dd"))
        'Set oDrNumDiaAnt = oCred.ObtieneNumeroRutaDiaAnt(Format(gdFecSis, "yyyy/MM/dd"))
    
        If Not (oDrNumDiaAnt.EOF And oDrNumDiaAnt.BOF) Then
            If (oDrNumDiaAnt!nCantidad > 0) Then
                MsgBox "Tiene pendiente registrar resultado de Hoja de Ruta Anterior", vbInformation, "AVISO"
                frmHojaRutaAnalistaResultado.Inicio 1
            End If
        End If
        If Not (oDrNumDiaAnt.EOF And oDrNumDiaAnt.BOF) Then
            If (oDrNumEnDia!nCantidad > 0) Then
            Else
                MsgBox "Debe Registrar su Hoja de Ruta de forma obligatoria", vbInformation, "AVISO"
                frmHojaRutaAnalista.Inicio 4, "Hoja de Ruta Analista"
            End If
        End If
    'FIN COMENTADO POR VAPI
    Else
        'AGREGADO POR VAPI ERS 0232015
        Screen.MousePointer = 0
        'Dim oDhoja As New DCOMhojaRuta
        Dim cPeriodo As String: cPeriodo = Format(gdFecSis, "YYYYMM")
        
'        If oDhoja.haConfiguradoAgencia(cPeriodo, gsCodAge) Then
'            Dim nPendiestesAtrasados As Integer: nPendiestesAtrasados = oDhoja.obtenerNumeroVisitasPendientes(gsCodUser, 0)
'            If nPendiestesAtrasados = 0 Then
'                Dim nPendiestes As Integer: nPendiestes = oDhoja.obtenerNumeroVisitasPendientes(gsCodUser, 1)
'                If nPendiestes > 0 Then
'                    If oDhoja.esHoraLimite Then
'                        frmHojaRutaAnalistaGeneraResultado.Show 1
'                    End If
'                Else
'                    If oDhoja.esHoraLimite Then
'                        frmHojaRutaAnalistaGenera.Inicio 1 'genera de mañana
'                    Else
'                        If oDhoja.obtenerNumeroVisitasRegistradasHoy(gsCodUser) <= 0 Then
'                            frmHojaRutaAnalistaGenera.Inicio 0 'genera de hoy
'                        End If
'                    End If
'                End If
'
'            Else
'                'ACA DEBE SOLICITAR EL VISTO DEL JEFE DE AGENCIA
'                If Not oDhoja.tieneVistoPendiente(gsCodUser) Then
'                    Dim cMovNro As String: cMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'                    oDhoja.solicitarVisto gsCodUser, cMovNro
'                End If
'                If frmHojaRutaAnalistaVistoInc.Inicio Then
'                    frmHojaRutaAnalistaGeneraResultado.Show 1
'                Else
'                    End
'                End If
'            End If
'        Else
'            MsgBox "Aún no existe Configuración de Hoja de Ruta para la Agencia, comuníquelo al Jefe de Agencia."
'        End If
    End If
    'FIN VAPI
End If
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
'****************APRI 20161024
    'INDICAR ACÁ LAS TECLAS
        'INDICAR ACÁ LAS TECLAS
     If RegisterHotKey(hwnd, 1, MOD_CONTROL, VK_F1) = 0 Or RegisterHotKey(hwnd, 1, MOD_CONTROL, VK_F2) = 0 Or RegisterHotKey(hwnd, 1, MOD_CONTROL, VK_F3) = 0 Or RegisterHotKey(hwnd, 1, MOD_CONTROL, VK_F4) = 0 Or RegisterHotKey(hwnd, 1, MOD_CONTROL, VK_F5) = 0 Or RegisterHotKey(hwnd, 1, MOD_CONTROL, VK_F6) = 0 Or RegisterHotKey(hwnd, 1, MOD_CONTROL, VK_F7) = 0 Or RegisterHotKey(hwnd, 1, MOD_CONTROL, VK_F8) = 0 Or RegisterHotKey(hwnd, 1, MOD_CONTROL, VK_F9) = 0 Or RegisterHotKey(hwnd, 1, MOD_CONTROL, VK_F10) = 0 Or RegisterHotKey(hwnd, 1, MOD_CONTROL, VK_F11) = 0 Or RegisterHotKey(hwnd, 1, MOD_CONTROL, VK_F12) = 0 Or RegisterHotKey(hwnd, 1, MOD_CONTROL, VK_1) = 0 Or RegisterHotKey(hwnd, 1, MOD_CONTROL, VK_2) = 0 Or RegisterHotKey(hwnd, 1, MOD_CONTROL, VK_3) = 0 Or RegisterHotKey(hwnd, 1, MOD_CONTROL, VK_4) = 0 Or RegisterHotKey(hwnd, 1, MOD_CONTROL, VK_5) = 0 Or RegisterHotKey(hwnd, 1, MOD_CONTROL, VK_6) = 0 Or RegisterHotKey(hwnd, 1, MOD_CONTROL, VK_7) = 0 Then
            'error
            MsgBox " Hubo un error ", vbCritical
        Exit Sub
      End If
      If RegisterHotKey(hwnd, 1, MOD_ALT, VK_1) = 0 Or RegisterHotKey(hwnd, 1, MOD_ALT, VK_2) = 0 Or RegisterHotKey(hwnd, 1, MOD_ALT, VK_3) = 0 Or RegisterHotKey(hwnd, 1, MOD_ALT, VK_4) = 0 Or RegisterHotKey(hwnd, 1, MOD_ALT, VK_5) = 0 Or RegisterHotKey(hwnd, 1, MOD_ALT, VK_6) = 0 Or RegisterHotKey(hwnd, 1, MOD_ALT, VK_7) = 0 Or RegisterHotKey(hwnd, 1, MOD_ALT, VK_8) = 0 Or RegisterHotKey(hwnd, 1, MOD_ALT, VK_9) = 0 Then
            'error
            MsgBox " Hubo un error ", vbCritical
        Exit Sub
      End If
'      WinProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf NewWindowProc)
'****************END APRI 20161024

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
        lsCadAux = lsCadAux + PstaNombre(Trim(lrst("cPersNombre")), True) + space(10) + Trim(lrst("cOpeDesOri")) + space(10) + Trim(lrst("sAutEstadoDes")) + Chr(13)
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
           frmCredCalendPagosNEW.InicioSim 'ARLO20180625
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
    'AppActivate valor             ' Activa la Calculadora. 'LUCV20190323, Comentó

End Sub


'**DAOR 20090203 , Registro de salida del sistema SICMACM Negocio************
'Bitacora Version 201011
Sub SalirSICMACMNegocio()
Dim oSeguridad As New COMManejador.Pista
    Call oSeguridad.InsertarPista(gIngresarSalirSistema, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gSalirSistema, "Salida del SICMACM Operaciones" & " Versión: " & Format(App.Major, "#0") & "." & Format(App.Minor, "#0") & "." & Format(App.Revision, "#0") & "-" & frmLogin.gsFechaVersion)
     If oSeguridad.ValidaAccesoPistaRF(gsCodUser) Then
            Call oSeguridad.InsertarPistaSesion(gIngresarSalirSistema, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gSalirSistema, 0)
            Call oSeguridad.ActualizarPistaSesion(gsCodPersUser, GetMaquinaUsuario, 0) 'JUEZ 20160125
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
    Dim J As Integer
            
    sGrupoAutorizado = oCons.LeeConstSistema(516)
    VerificaGrupoPermisoPostCierre = False
    For i = 1 To Len(sGrupoAutorizado)
        If Not Mid(sGrupoAutorizado, i, 1) = "," Then
            nGrupoTmp1 = nGrupoTmp1 & Mid(sGrupoAutorizado, i, 1)
        Else
            For J = 1 To Len(gsGruposUser)
                If Not Mid(gsGruposUser, J, 1) = "," Then
                    nGrupoTmp2 = nGrupoTmp2 & Mid(gsGruposUser, J, 1)
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
    Dim J As Integer
            
    sGrupoAutorizado = oCons.LeeConstSistema(519)
    VerificaGrupoMantenimientoUsuarios = False
    For i = 1 To Len(sGrupoAutorizado)
        If Not Mid(sGrupoAutorizado, i, 1) = "," Then
            nGrupoTmp1 = nGrupoTmp1 & Mid(sGrupoAutorizado, i, 1)
        Else
            For J = 1 To Len(gsGruposUser)
                If Not Mid(gsGruposUser, J, 1) = "," Then
                    nGrupoTmp2 = nGrupoTmp2 & Mid(gsGruposUser, J, 1)
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
    Image1.Move Pos_x, Pos_y - 180
End Sub
'<-***** Fin LUCV20190323
