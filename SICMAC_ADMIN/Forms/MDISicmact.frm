VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.MDIForm MDISicmact 
   BackColor       =   &H8000000C&
   Caption         =   "SICMACT Sistema Integrado de la Caja Municipal de Ahorro y Credito de Trujillo"
   ClientHeight    =   8880
   ClientLeft      =   750
   ClientTop       =   525
   ClientWidth     =   14700
   Icon            =   "MDISicmact.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pbxFondo 
      Align           =   1  'Align Top
      Height          =   7935
      Left            =   0
      Picture         =   "MDISicmact.frx":030A
      ScaleHeight     =   7875
      ScaleWidth      =   14640
      TabIndex        =   2
      Top             =   600
      Width           =   14700
      Begin VB.Image Image1 
         Height          =   13500
         Left            =   0
         Picture         =   "MDISicmact.frx":51586
         Top             =   0
         Width           =   21600
      End
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14700
      _ExtentX        =   25929
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlMain"
      HotImageList    =   "imlMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Calculadora"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mantenimiento Permisos"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlMain 
      Left            =   600
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDISicmact.frx":7843F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDISicmact.frx":79491
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staMain 
      Align           =   2  'Align Bottom
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   8655
      Width           =   14700
      _ExtentX        =   25929
      _ExtentY        =   397
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Timer Tiempo 
      Interval        =   60000
      Left            =   120
      Top             =   840
   End
   Begin VB.Menu M0100000000 
      Caption         =   "Arc&hivo"
      Index           =   0
      Begin VB.Menu M0101000000 
         Caption         =   "Configurar &Impresora"
         Index           =   0
      End
      Begin VB.Menu M0101000000 
         Caption         =   "Configurar Caracteres de Impresión"
         Index           =   1
      End
      Begin VB.Menu M0101000000 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu M0101000000 
         Caption         =   "&Salir"
         Index           =   3
      End
   End
   Begin VB.Menu M0700000000 
      Caption         =   "Perso&nas"
      Index           =   0
      Begin VB.Menu M0701000000 
         Caption         =   "&Personas"
         Index           =   0
         Begin VB.Menu M0701010000 
            Caption         =   "&Consulta"
            Index           =   2
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
         Caption         =   "Grupos Economicos"
         Index           =   2
         Begin VB.Menu M0701030000 
            Caption         =   "Mantenimiento"
            Index           =   0
         End
      End
      Begin VB.Menu M0701000000 
         Caption         =   "Promociones"
         Index           =   3
         Visible         =   0   'False
         Begin VB.Menu M0701040000 
            Caption         =   "Registro"
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu M0701040000 
            Caption         =   "Reportes"
            Index           =   1
         End
      End
   End
   Begin VB.Menu M0800000000 
      Caption         =   "Cli&ente"
      Index           =   0
      Begin VB.Menu M0801000000 
         Caption         =   "&Posicion de Cliente"
         Index           =   0
      End
   End
   Begin VB.Menu M0900000000 
      Caption         =   "Herra&mientas"
      Index           =   0
      Begin VB.Menu M0901000000 
         Caption         =   "Editor de Textos"
         Index           =   0
      End
      Begin VB.Menu M0901000000 
         Caption         =   "&Spooler de Impresion"
         Index           =   1
      End
      Begin VB.Menu M0901000000 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu M0901000000 
         Caption         =   "Configuracion &Perifericos"
         Index           =   3
      End
      Begin VB.Menu M0901000000 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu M0901000000 
         Caption         =   "BackUp"
         Index           =   5
      End
   End
   Begin VB.Menu M1000000000 
      Caption         =   "&Seguridad"
      Index           =   0
      Begin VB.Menu M1001000000 
         Caption         =   "Asignar Permisos"
         Index           =   0
      End
      Begin VB.Menu M1001000000 
         Caption         =   "Administracion de Usuarios"
         Index           =   1
      End
   End
   Begin VB.Menu M1600000000 
      Caption         =   "&Recursos Humanos"
      Index           =   0
      Begin VB.Menu M1601000000 
         Caption         =   "&Selección"
         Index           =   0
         Begin VB.Menu M1601010000 
            Caption         =   "&Inicio Proceso Selección"
            Index           =   0
            Begin VB.Menu M1601010100 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M1601010100 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M1601010100 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu M1601010100 
               Caption         =   "&Reporte"
               Index           =   3
            End
         End
         Begin VB.Menu M1601010000 
            Caption         =   "&Postulantes"
            Index           =   1
            Begin VB.Menu M1601010200 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M1601010200 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M1601010200 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu M1601010200 
               Caption         =   "&Reporte"
               Index           =   3
            End
         End
         Begin VB.Menu M1601010000 
            Caption         =   "&Evaluación"
            Index           =   2
            Begin VB.Menu M1601010300 
               Caption         =   "&Curricular"
               Index           =   0
               Begin VB.Menu M1601010301 
                  Caption         =   "&Registro"
                  Index           =   0
               End
               Begin VB.Menu M1601010301 
                  Caption         =   "&Mantenimiento"
                  Index           =   1
               End
               Begin VB.Menu M1601010301 
                  Caption         =   "&Consulta"
                  Index           =   2
               End
               Begin VB.Menu M1601010301 
                  Caption         =   "&Reporte"
                  Index           =   3
               End
            End
            Begin VB.Menu M1601010300 
               Caption         =   "&Escrita"
               Index           =   1
               Begin VB.Menu M1601010302 
                  Caption         =   "&Registro"
                  Index           =   0
               End
               Begin VB.Menu M1601010302 
                  Caption         =   "&Mantenimiento"
                  Index           =   1
               End
               Begin VB.Menu M1601010302 
                  Caption         =   "&Consulta"
                  Index           =   2
               End
               Begin VB.Menu M1601010302 
                  Caption         =   "&Reporte"
                  Index           =   3
               End
            End
            Begin VB.Menu M1601010300 
               Caption         =   "&Psicologica"
               Index           =   2
               Begin VB.Menu M1601010303 
                  Caption         =   "&Registro"
                  Index           =   0
               End
               Begin VB.Menu M1601010303 
                  Caption         =   "&Mantenimiento"
                  Index           =   1
               End
               Begin VB.Menu M1601010303 
                  Caption         =   "&Consulta"
                  Index           =   2
               End
               Begin VB.Menu M1601010303 
                  Caption         =   "&Reporte"
                  Index           =   3
               End
            End
            Begin VB.Menu M1601010300 
               Caption         =   "E&ntrevista"
               Index           =   3
               Begin VB.Menu M1601010304 
                  Caption         =   "&Registro"
                  Index           =   0
               End
               Begin VB.Menu M1601010304 
                  Caption         =   "&Mantenimiento"
                  Index           =   1
               End
               Begin VB.Menu M1601010304 
                  Caption         =   "&Consulta"
                  Index           =   2
               End
               Begin VB.Menu M1601010304 
                  Caption         =   "&Reporte"
                  Index           =   3
               End
            End
         End
         Begin VB.Menu M1601010000 
            Caption         =   "&Resultados y Cierre"
            Index           =   3
         End
         Begin VB.Menu M1601010000 
            Caption         =   "&Consulta"
            Index           =   4
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu M1601000000 
         Caption         =   "&Contratos"
         Index           =   2
         Begin VB.Menu M1601020000 
            Caption         =   "&Registro basado en el Proceso de Selección"
            Index           =   0
         End
         Begin VB.Menu M1601020000 
            Caption         =   "R&egistro manual"
            Index           =   1
         End
         Begin VB.Menu M1601020000 
            Caption         =   "&Mantenimiento"
            Index           =   2
         End
         Begin VB.Menu M1601020000 
            Caption         =   "&Consulta"
            Index           =   3
         End
         Begin VB.Menu M1601020000 
            Caption         =   "Re&scindir"
            Index           =   4
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu M1601000000 
         Caption         =   "&Adenda"
         Index           =   4
         Begin VB.Menu M1601030000 
            Caption         =   "&Registro"
            Index           =   0
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu M1601000000 
         Caption         =   "C&urriculum Vitae"
         Index           =   6
         Begin VB.Menu M1601040000 
            Caption         =   "&Configuracion"
            Index           =   0
            Begin VB.Menu M1601040100 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M1601040100 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M1601040000 
            Caption         =   "&Registro"
            Index           =   1
         End
         Begin VB.Menu M1601040000 
            Caption         =   "&Mantenimiento"
            Index           =   2
         End
         Begin VB.Menu M1601040000 
            Caption         =   "&Consulta"
            Index           =   3
         End
         Begin VB.Menu M1601040000 
            Caption         =   "&Reporte"
            Index           =   4
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu M1601000000 
         Caption         =   "Ho&rario Laboral"
         Index           =   8
         Begin VB.Menu M1601050000 
            Caption         =   "&Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu M1601050000 
            Caption         =   "&Consulta"
            Index           =   1
         End
         Begin VB.Menu M1601050000 
            Caption         =   "&Horarios"
            Index           =   2
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "&Asistencia"
         Index           =   9
      End
      Begin VB.Menu M1601000000 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu M1601000000 
         Caption         =   "E&valuacion Interna"
         Index           =   11
         Begin VB.Menu M1601060000 
            Caption         =   "&Inicio de proceso de Evaluacion Interna"
            Index           =   0
            Begin VB.Menu M1601060100 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M1601060100 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M1601060100 
               Caption         =   "&Consulta"
               Index           =   2
            End
         End
         Begin VB.Menu M1601060000 
            Caption         =   "&Curricular"
            Index           =   1
            Begin VB.Menu M1601060200 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M1601060200 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M1601060200 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu M1601060200 
               Caption         =   "&Reporte"
               Index           =   3
            End
         End
         Begin VB.Menu M1601060000 
            Caption         =   "&Escrita"
            Index           =   2
            Begin VB.Menu M1601060300 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M1601060300 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M1601060300 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu M1601060300 
               Caption         =   "&Reporte"
               Index           =   3
            End
         End
         Begin VB.Menu M1601060000 
            Caption         =   "&Psicologica"
            Index           =   3
            Begin VB.Menu M1601060400 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M1601060400 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M1601060400 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu M1601060400 
               Caption         =   "&Reporte"
               Index           =   3
            End
         End
         Begin VB.Menu M1601060000 
            Caption         =   "E&ntrevista"
            Index           =   4
            Begin VB.Menu M1601060500 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M1601060500 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M1601060500 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu M1601060500 
               Caption         =   "&Reporte"
               Index           =   3
            End
         End
         Begin VB.Menu M1601060000 
            Caption         =   "Resultados y Cie&rre"
            Index           =   5
         End
         Begin VB.Menu M1601060000 
            Caption         =   "&Consulta"
            Index           =   6
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu M1601000000 
         Caption         =   "&Permisos"
         Index           =   13
         Begin VB.Menu M1601070000 
            Caption         =   "&Solicitud"
            Index           =   0
            Begin VB.Menu M1601070100 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M1601070100 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M1601070000 
            Caption         =   "&Autorizacion/Rechazo"
            Index           =   1
         End
         Begin VB.Menu M1601070000 
            Caption         =   "&Reporte"
            Index           =   2
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "&Vacaciones"
         Index           =   14
         Begin VB.Menu M1601080000 
            Caption         =   "&Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu M1601080000 
            Caption         =   "&Consulta"
            Index           =   1
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "&Descansos/Subsidios"
         Index           =   15
         Begin VB.Menu M1601090000 
            Caption         =   "&Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu M1601090000 
            Caption         =   "&Consulta"
            Index           =   1
         End
         Begin VB.Menu M1601090000 
            Caption         =   "&Reportes"
            Index           =   2
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "&Suspensiones"
         Index           =   16
         Begin VB.Menu M1601100000 
            Caption         =   "&Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu M1601100000 
            Caption         =   "&Consulta"
            Index           =   1
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu M1601000000 
         Caption         =   "&Meritos y Demeritos"
         Index           =   18
         Begin VB.Menu M1601110000 
            Caption         =   "&Tabla"
            Index           =   0
            Begin VB.Menu M1601110100 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M1601110100 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M1601110000 
            Caption         =   "&Mantenimiento"
            Index           =   1
         End
         Begin VB.Menu M1601110000 
            Caption         =   "&Consulta"
            Index           =   2
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "Car&gos Laborales"
         Index           =   19
         Begin VB.Menu M1601120000 
            Caption         =   "&Tabla"
            Index           =   0
            Begin VB.Menu M1601120100 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M1601120100 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M1601120000 
            Caption         =   "&Registro"
            Index           =   1
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "Sueldos"
         Index           =   20
         Begin VB.Menu M1601130000 
            Caption         =   "&Registro"
            Index           =   0
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "Sistema de Pensiones"
         Index           =   21
         Begin VB.Menu M1601140000 
            Caption         =   "&Tabla"
            Index           =   0
            Begin VB.Menu M1601140100 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
         End
         Begin VB.Menu M1601140000 
            Caption         =   "&Registro"
            Index           =   1
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "InformeSocial"
         Index           =   22
         Begin VB.Menu M1601150000 
            Caption         =   "&Registro"
            Index           =   0
         End
         Begin VB.Menu M1601150000 
            Caption         =   "&Mantenimiento"
            Index           =   1
         End
         Begin VB.Menu M1601150000 
            Caption         =   "&Consulta"
            Index           =   2
         End
         Begin VB.Menu M1601150000 
            Caption         =   "&Reportes"
            Index           =   3
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "Comentario"
         Index           =   23
      End
      Begin VB.Menu M1601000000 
         Caption         =   "Asistencia Medica"
         Index           =   24
         Begin VB.Menu M1601160000 
            Caption         =   "&Tabla"
            Index           =   0
            Begin VB.Menu M1601160100 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
         End
         Begin VB.Menu M1601160000 
            Caption         =   "&Registro"
            Index           =   1
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "-"
         Index           =   25
      End
      Begin VB.Menu M1601000000 
         Caption         =   "Confi&guracion de Planilla de Remuneraciones"
         Index           =   26
         Begin VB.Menu M1601170000 
            Caption         =   "&Configuracion de Conceptos Remunerativos"
            Index           =   0
            Begin VB.Menu M1601170100 
               Caption         =   "&Mantenimiento de Conceptos Remunerativos"
               Index           =   0
            End
            Begin VB.Menu M1601170100 
               Caption         =   "&Mantenimiento de &Tabla Alias"
               Index           =   1
            End
            Begin VB.Menu M1601170100 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu M1601170100 
               Caption         =   "&Reporte"
               Index           =   3
            End
            Begin VB.Menu M1601170100 
               Caption         =   "Cuentas Con&tables"
               Index           =   4
            End
         End
         Begin VB.Menu M1601170000 
            Caption         =   "C&onfiguracion de Conceptos de Planilla de Remuneraciones"
            Index           =   1
            Begin VB.Menu M1601170200 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M1601170200 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M1601170200 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu M1601170200 
               Caption         =   "&Reporte"
               Index           =   3
            End
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "&Planilla de remuneraciones del RR.HH."
         Index           =   27
         Begin VB.Menu M1601180000 
            Caption         =   "&Conceptos Fijos de Planilla de Remuneraciones"
            Index           =   0
            Begin VB.Menu M1601180100 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M1601180100 
               Caption         =   "&Consulta"
               Index           =   1
            End
            Begin VB.Menu M1601180100 
               Caption         =   "&Reporte"
               Index           =   2
            End
            Begin VB.Menu M1601180100 
               Caption         =   "&Pre Planilla "
               Index           =   3
            End
         End
         Begin VB.Menu M1601180000 
            Caption         =   "Conceptos &Variables de Planilla de Remuneraciones"
            Index           =   1
            Begin VB.Menu M1601180200 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M1601180200 
               Caption         =   "&Consulta"
               Index           =   1
            End
            Begin VB.Menu M1601180200 
               Caption         =   "&Reporte"
               Index           =   2
            End
         End
         Begin VB.Menu M1601180000 
            Caption         =   "E&xtra Planilla Variables (Ingresos y Descuentos)"
            Index           =   2
            Begin VB.Menu M1601180300 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M1601180300 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M1601180300 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu M1601180300 
               Caption         =   "&Reporte"
               Index           =   3
            End
            Begin VB.Menu M1601180300 
               Caption         =   "-"
               Index           =   4
            End
            Begin VB.Menu M1601180300 
               Caption         =   "Cuentas de Abono No Empleados"
               Index           =   5
            End
         End
         Begin VB.Menu M1601180000 
            Caption         =   "Creditos &Administrativos"
            Index           =   3
            Begin VB.Menu M1601180400 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M1601180400 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M1601180000 
            Caption         =   "Creditos en &Otras Entidades"
            Index           =   4
            Begin VB.Menu M1601180500 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M1601180500 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "-"
         Index           =   28
      End
      Begin VB.Menu M1601000000 
         Caption         =   "P&rocesos"
         Index           =   29
         Begin VB.Menu M1601190000 
            Caption         =   "&Calculo de Planillas"
            Index           =   0
         End
         Begin VB.Menu M1601190000 
            Caption         =   "&Abono a Cuenta"
            Index           =   1
         End
         Begin VB.Menu M1601190000 
            Caption         =   "Cierre Mensual"
            Index           =   2
         End
         Begin VB.Menu M1601190000 
            Caption         =   "Cierre &Diario"
            Index           =   3
         End
         Begin VB.Menu M1601190000 
            Caption         =   "Pago Asistencia Medica P&rivada"
            Enabled         =   0   'False
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu M1601190000 
            Caption         =   "Periodos No Laborales"
            Index           =   5
         End
         Begin VB.Menu M1601190000 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu M1601190001 
            Caption         =   "Provisiones"
            Index           =   0
            Begin VB.Menu M1601190002 
               Caption         =   "Vacaciones"
               Index           =   0
            End
            Begin VB.Menu M1601190002 
               Caption         =   "CTS"
               Index           =   1
            End
            Begin VB.Menu M1601190002 
               Caption         =   "Gratificaciones"
               Index           =   2
            End
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "-"
         Index           =   30
      End
      Begin VB.Menu M1601000000 
         Caption         =   "&Reportes"
         Index           =   31
         Begin VB.Menu M1601200000 
            Caption         =   "&Reportes"
            Index           =   0
         End
         Begin VB.Menu M1601200000 
            Caption         =   "Reportes &Generales"
            Index           =   1
         End
         Begin VB.Menu M1601200000 
            Caption         =   "Cargo Actual"
            Index           =   2
         End
         Begin VB.Menu M1601200000 
            Caption         =   "Expediente RRHH"
            Index           =   3
         End
      End
      Begin VB.Menu M1601000000 
         Caption         =   "Registrar Expediente"
         Index           =   32
      End
      Begin VB.Menu M1601000000 
         Caption         =   "Registrar Expediente Duración"
         Index           =   33
      End
   End
   Begin VB.Menu M1700000000 
      Caption         =   "&Logística"
      Index           =   0
      Begin VB.Menu M1701000000 
         Caption         =   "&Bienes"
         Index           =   2
      End
      Begin VB.Menu M1701000000 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu M1701000000 
         Caption         =   "&Proveedores"
         Index           =   4
      End
      Begin VB.Menu M1701000000 
         Caption         =   "-"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Requerimiento Regular"
         Index           =   6
         Visible         =   0   'False
         Begin VB.Menu M1701010000 
            Caption         =   "&Requerimientos"
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu M1701010000 
            Caption         =   "&Aprobación de Requerimiento"
            Index           =   1
         End
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Plan Anual de Adquisiciones y Contrataciones"
         Index           =   9
         Visible         =   0   'False
         Begin VB.Menu M1701030000 
            Caption         =   "Consolidación de los Requerimientos "
            Index           =   0
         End
         Begin VB.Menu M1701030000 
            Caption         =   "Plan Anual de Adqusiciones Segun la  Consolidacion "
            Index           =   1
         End
      End
      Begin VB.Menu M1701000000 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Proceso de Selección"
         Index           =   11
         Visible         =   0   'False
         Begin VB.Menu M1701040000 
            Caption         =   "Comite"
            Index           =   0
         End
         Begin VB.Menu M1701040000 
            Caption         =   "Proceso de Seleccion"
            Index           =   1
         End
         Begin VB.Menu M1701040000 
            Caption         =   "Proveedores"
            Index           =   2
         End
         Begin VB.Menu M1701040000 
            Caption         =   "Configuracion de Bienes"
            Index           =   3
         End
         Begin VB.Menu M1701040000 
            Caption         =   "Criterios de Evaluacion"
            Index           =   4
            Begin VB.Menu M1701040300 
               Caption         =   "Mantenimiento de Criterios"
               Index           =   0
            End
            Begin VB.Menu M1701040300 
               Caption         =   "Asignacion de Criterios a Proceso"
               Index           =   1
            End
         End
         Begin VB.Menu M1701040000 
            Caption         =   "Cotizacion"
            Index           =   5
            Begin VB.Menu M1701040400 
               Caption         =   "Solicitud de Cotizacion"
               Index           =   0
            End
         End
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Evaluacion de Proceso de Seleccion"
         Index           =   12
         Visible         =   0   'False
         Begin VB.Menu M1701040600 
            Caption         =   "Evaluacion Tecnica"
            Index           =   0
            Begin VB.Menu M1701040650 
               Caption         =   "Registro"
               Index           =   0
            End
            Begin VB.Menu M1701040650 
               Caption         =   "Resumen"
               Index           =   1
            End
         End
         Begin VB.Menu M1701040600 
            Caption         =   "Evaluacion Economica"
            Index           =   1
            Begin VB.Menu M1701040660 
               Caption         =   "Registro Cotizacion"
               Index           =   0
            End
            Begin VB.Menu M1701040660 
               Caption         =   "Resumen"
               Index           =   1
            End
         End
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Cuadro Comparativo de Cotizaciones"
         Index           =   13
         Visible         =   0   'False
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Cancelacion de Proceso de Seleccion"
         Index           =   14
         Visible         =   0   'False
      End
      Begin VB.Menu M1701000000 
         Caption         =   "-"
         Index           =   17
         Visible         =   0   'False
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Contratación"
         Index           =   18
         Begin VB.Menu M1701050000 
            Caption         =   "&Solicitud"
            Index           =   0
            Begin VB.Menu M1701050100 
               Caption         =   "Orden Compra Soles"
               Index           =   0
            End
            Begin VB.Menu M1701050100 
               Caption         =   "Orden Compra Dolares"
               Index           =   1
            End
            Begin VB.Menu M1701050100 
               Caption         =   "Orden Servicio Soles"
               Index           =   2
            End
            Begin VB.Menu M1701050100 
               Caption         =   "Orden Servicio Dolares"
               Index           =   3
            End
         End
         Begin VB.Menu M1701050000 
            Caption         =   "&Mantenimiento Contratación"
            Index           =   1
            Begin VB.Menu M1701050200 
               Caption         =   "Orden Compra Soles"
               Index           =   0
            End
            Begin VB.Menu M1701050200 
               Caption         =   "Orden Compra Dolares"
               Index           =   1
            End
            Begin VB.Menu M1701050200 
               Caption         =   "Orden Servicio Soles"
               Index           =   2
            End
            Begin VB.Menu M1701050200 
               Caption         =   "Orden Servicio Dolares"
               Index           =   3
            End
            Begin VB.Menu M1701050200 
               Caption         =   "Seguimiento de Ordenes"
               Index           =   4
            End
         End
         Begin VB.Menu M1701050000 
            Caption         =   "&Impresión de Contratación"
            Index           =   2
            Begin VB.Menu M1701050300 
               Caption         =   "Orden Compra Soles"
               Index           =   0
            End
            Begin VB.Menu M1701050300 
               Caption         =   "Orden Compra Dolares"
               Index           =   1
            End
            Begin VB.Menu M1701050300 
               Caption         =   "Orden Servicio Soles"
               Index           =   2
            End
            Begin VB.Menu M1701050300 
               Caption         =   "Orden Servicio Dolares"
               Index           =   3
            End
         End
         Begin VB.Menu M1701050000 
            Caption         =   "Impresion por Fechas"
            Index           =   3
            Begin VB.Menu M1701050400 
               Caption         =   "Orden Compra Soles"
               Index           =   0
            End
            Begin VB.Menu M1701050400 
               Caption         =   "Orden Compra Dolares"
               Index           =   1
            End
            Begin VB.Menu M1701050400 
               Caption         =   "Orden Servicio Soles"
               Index           =   2
            End
            Begin VB.Menu M1701050400 
               Caption         =   "Orden Servicio Dolares"
               Index           =   3
            End
         End
         Begin VB.Menu M1701050000 
            Caption         =   "Contratos"
            Index           =   4
            Begin VB.Menu M1701050500 
               Caption         =   "Registro de Contratos"
               Index           =   0
            End
            Begin VB.Menu M1701050500 
               Caption         =   "Seguimiento de Contratos"
               Index           =   1
            End
         End
         Begin VB.Menu M1701050000 
            Caption         =   "Registro de Comprobantes"
            Enabled         =   0   'False
            Index           =   5
            Visible         =   0   'False
            Begin VB.Menu M1701050600 
               Caption         =   "Órdenes de Compra Soles"
               Enabled         =   0   'False
               Index           =   0
            End
            Begin VB.Menu M1701050600 
               Caption         =   "Órdenes de Compra Dólares"
               Enabled         =   0   'False
               Index           =   1
               Visible         =   0   'False
            End
            Begin VB.Menu M1701050600 
               Caption         =   "Órdenes de Servicio Soles"
               Enabled         =   0   'False
               Index           =   2
               Visible         =   0   'False
            End
            Begin VB.Menu M1701050600 
               Caption         =   "Órdenes de Servicio Dólares"
               Enabled         =   0   'False
               Index           =   3
               Visible         =   0   'False
            End
            Begin VB.Menu M1701050600 
               Caption         =   "Contratos"
               Enabled         =   0   'False
               Index           =   4
               Visible         =   0   'False
            End
            Begin VB.Menu M1701050600 
               Caption         =   "Adendas"
               Enabled         =   0   'False
               Index           =   5
               Visible         =   0   'False
            End
            Begin VB.Menu M1701050600 
               Caption         =   "Compras Directas"
               Enabled         =   0   'False
               Index           =   6
               Visible         =   0   'False
            End
            Begin VB.Menu M1701050600 
               Caption         =   "Impresion de Comprobantes"
               Enabled         =   0   'False
               Index           =   7
               Visible         =   0   'False
            End
         End
         Begin VB.Menu M1701050000 
            Caption         =   "Acta de Conformidad"
            Index           =   6
            Begin VB.Menu M1701050700 
               Caption         =   "Acta de Conformidad Digital Soles"
               Index           =   0
            End
            Begin VB.Menu M1701050700 
               Caption         =   "Acta de Conformidad Digital Dolares"
               Index           =   1
            End
            Begin VB.Menu M1701050700 
               Caption         =   "Acta de Conformidad Digital Libre Soles"
               Index           =   2
            End
            Begin VB.Menu M1701050700 
               Caption         =   "Acta de Conformidad Digital Libre Dolares"
               Index           =   3
            End
            Begin VB.Menu M1701050700 
               Caption         =   "Extorno de Acta de Conformidad"
               Index           =   4
            End
         End
         Begin VB.Menu M1701050000 
            Caption         =   "Comprobantes"
            Index           =   7
            Begin VB.Menu M1701050800 
               Caption         =   "Registro"
               Enabled         =   0   'False
               Index           =   0
            End
            Begin VB.Menu M1701050800 
               Caption         =   "Registro"
               Index           =   1
               Begin VB.Menu M1701050801 
                  Caption         =   "Comprobante Soles"
                  Index           =   0
               End
               Begin VB.Menu M1701050801 
                  Caption         =   "Comprobante Dólares"
                  Index           =   1
               End
               Begin VB.Menu M1701050801 
                  Caption         =   "Comprobante Libre Soles"
                  Index           =   2
               End
               Begin VB.Menu M1701050801 
                  Caption         =   "Comprobante Libre Dólares"
                  Index           =   3
               End
            End
            Begin VB.Menu M1701050800 
               Caption         =   "Impresión Comprobante"
               Index           =   2
            End
            Begin VB.Menu M1701050800 
               Caption         =   "Historial de Comprobantes"
               Index           =   3
            End
            Begin VB.Menu M1701050800 
               Caption         =   "Extorno"
               Index           =   4
               Begin VB.Menu M1701050804 
                  Caption         =   "Comprobante Soles"
                  Index           =   0
               End
               Begin VB.Menu M1701050804 
                  Caption         =   "Comprobante Dólares"
                  Index           =   1
               End
               Begin VB.Menu M1701050804 
                  Caption         =   "Comprobante Libre Soles"
                  Index           =   2
               End
               Begin VB.Menu M1701050804 
                  Caption         =   "Comprobante Libre Dólares"
                  Index           =   3
               End
            End
         End
         Begin VB.Menu M1701050000 
            Caption         =   "Carta Fianza"
            Index           =   8
         End
      End
      Begin VB.Menu M1701000000 
         Caption         =   "-"
         Index           =   20
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Almacén"
         Index           =   21
         Begin VB.Menu M1701070000 
            Caption         =   "Operaciones"
            Index           =   0
         End
         Begin VB.Menu M1701070000 
            Caption         =   "Inventario"
            Index           =   1
         End
         Begin VB.Menu M1701070000 
            Caption         =   "Kardex"
            Index           =   2
         End
         Begin VB.Menu M1701070000 
            Caption         =   "Cierre de Saldos x Agencia"
            Index           =   3
         End
         Begin VB.Menu M1701070000 
            Caption         =   "Definicon de Cta Contables"
            Index           =   4
         End
         Begin VB.Menu M1701070000 
            Caption         =   "Mantenimiento de Saldos"
            Index           =   5
         End
         Begin VB.Menu M1701070000 
            Caption         =   "Regeneraciòn de Asientos Contables"
            Enabled         =   0   'False
            Index           =   6
         End
         Begin VB.Menu M1701070000 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu M1701070000 
            Caption         =   "Reportes Estadisticos"
            Index           =   10
         End
         Begin VB.Menu M1701070000 
            Caption         =   "Estadisticas de Consumo"
            Index           =   11
         End
         Begin VB.Menu M1701070000 
            Caption         =   "Estadisticas de Atencion"
            Index           =   12
         End
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Activo Fijo"
         Index           =   22
         Begin VB.Menu M1701080000 
            Caption         =   "Depreciación de Bienes"
            Index           =   0
         End
         Begin VB.Menu M1701080000 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu M1701080000 
            Caption         =   "Transferencia de Activos Fijos y Bienes"
            Index           =   2
         End
         Begin VB.Menu M1701080000 
            Caption         =   "Asignacion de Activo Fijo"
            Index           =   3
         End
         Begin VB.Menu M1701080000 
            Caption         =   "Baja de Activos"
            Index           =   4
         End
         Begin VB.Menu M1701080000 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu M1701080000 
            Caption         =   "Reporte de Bienes"
            Index           =   6
         End
         Begin VB.Menu M1701080000 
            Caption         =   "Kardex de Activo"
            Index           =   7
         End
         Begin VB.Menu M1701080000 
            Caption         =   "Modificación de Activos Fijos y Bienes"
            Index           =   8
         End
         Begin VB.Menu M1701080000 
            Caption         =   "Ajuste de Vida Útil"
            Index           =   9
         End
         Begin VB.Menu M1701080000 
            Caption         =   "Destino Bienes Dados de Baja"
            Index           =   10
         End
         Begin VB.Menu M1701080000 
            Caption         =   "Deterioro de Activos Fijos"
            Index           =   11
         End
      End
      Begin VB.Menu M1701000000 
         Caption         =   "-"
         Index           =   25
         Visible         =   0   'False
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Servicios (locación, públicos, privados, móviles, otros)"
         Index           =   26
         Visible         =   0   'False
         Begin VB.Menu M1701100000 
            Caption         =   "Registro"
            Index           =   0
         End
         Begin VB.Menu M1701100000 
            Caption         =   "Distribución"
            Index           =   1
         End
         Begin VB.Menu M1701100000 
            Caption         =   "Garantías"
            Index           =   2
         End
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Valorización de Inmuebles"
         Index           =   27
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Saneamiento"
         Index           =   28
      End
      Begin VB.Menu M1701000000 
         Caption         =   "-"
         Index           =   29
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Tasación Bienes Adjudicados"
         Index           =   30
      End
   End
   Begin VB.Menu M1800000000 
      Caption         =   "Presupuesto"
      Index           =   0
      Begin VB.Menu M1801000000 
         Caption         =   "Mantenimiento de Presupuesto"
         Index           =   0
      End
      Begin VB.Menu M1801000000 
         Caption         =   "Ingreso de Rubros"
         Index           =   1
      End
      Begin VB.Menu M1801000000 
         Caption         =   "Ingreso de Presupuesto"
         Index           =   2
      End
      Begin VB.Menu M1801000000 
         Caption         =   "Ejecución de Presupuesto"
         Index           =   3
      End
      Begin VB.Menu M1801000000 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu M1801000000 
         Caption         =   "Bienes"
         Index           =   5
         Begin VB.Menu M1801010000 
            Caption         =   "Soles"
            Index           =   0
         End
         Begin VB.Menu M1801010000 
            Caption         =   "Dolares"
            Index           =   1
         End
      End
      Begin VB.Menu M1801000000 
         Caption         =   "Servicios"
         Index           =   6
         Begin VB.Menu M1801020000 
            Caption         =   "Soles"
            Index           =   0
         End
         Begin VB.Menu M1801020000 
            Caption         =   "Dolares"
            Index           =   1
         End
      End
   End
   Begin VB.Menu M1900000000 
      Caption         =   "SI&G"
      Index           =   0
      Begin VB.Menu M1901000000 
         Caption         =   "&Recursos Humanos"
         Index           =   0
         Begin VB.Menu M1901010000 
            Caption         =   "Proceso de &Selección"
            Index           =   0
         End
         Begin VB.Menu M1901010000 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu M1901010000 
            Caption         =   "Recurso Humano"
            Index           =   2
         End
         Begin VB.Menu M1901010000 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu M1901010000 
            Caption         =   "&Asistencia "
            Index           =   4
         End
         Begin VB.Menu M1901010000 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu M1901010000 
            Caption         =   "&Planillas"
            Index           =   6
         End
         Begin VB.Menu M1901010000 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu M1901010000 
            Caption         =   "&Reportes"
            Index           =   8
         End
         Begin VB.Menu M1901010000 
            Caption         =   "Reportes &Generales"
            Index           =   9
         End
         Begin VB.Menu M1901010000 
            Caption         =   "Reporte Presupuesto"
            Index           =   10
         End
      End
      Begin VB.Menu M1901000000 
         Caption         =   "&Logistica"
         Index           =   1
         Begin VB.Menu M1901020000 
            Caption         =   "Saldos &Mensuales"
            Index           =   0
         End
         Begin VB.Menu M1901020000 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu M1901020000 
            Caption         =   "&Reportes"
            Index           =   2
         End
      End
      Begin VB.Menu M1901000000 
         Caption         =   "&Presupuesto"
         Index           =   2
         Begin VB.Menu M1901030000 
            Caption         =   "&Ejecucion"
            Index           =   0
         End
      End
   End
   Begin VB.Menu M2000000000 
      Caption         =   "Inventario"
      Index           =   0
      Begin VB.Menu M2001000000 
         Caption         =   "Activar Bienes"
         Index           =   0
         Begin VB.Menu M2001010000 
            Caption         =   "Soles"
            Index           =   1
         End
         Begin VB.Menu M2001020000 
            Caption         =   "Dolares"
            Index           =   2
         End
      End
      Begin VB.Menu M2002000000 
         Caption         =   "Transferencias"
         Index           =   3
      End
      Begin VB.Menu M2003000000 
         Caption         =   "Reportes"
         Index           =   4
         Begin VB.Menu M2003010000 
            Caption         =   "Activos Fijos"
            Index           =   5
         End
         Begin VB.Menu M2003020000 
            Caption         =   "Asiento de las Transferencias"
            Index           =   5
         End
         Begin VB.Menu M2003030000 
            Caption         =   "Transferencia"
            Index           =   6
         End
      End
   End
   Begin VB.Menu M2100000000 
      Caption         =   "Marketing"
      Index           =   0
      Begin VB.Menu M2101000000 
         Caption         =   "Parámetros de Gastos"
         Index           =   0
         Begin VB.Menu M2101010000 
            Caption         =   "Registro"
            Index           =   0
         End
         Begin VB.Menu M2101010000 
            Caption         =   "Mantenimiento"
            Index           =   1
         End
         Begin VB.Menu M2101010000 
            Caption         =   "Consulta"
            Index           =   2
         End
      End
      Begin VB.Menu M2101000000 
         Caption         =   "Registro de Productos y Servicios"
         Index           =   1
      End
      Begin VB.Menu M2101000000 
         Caption         =   "Registro de Actividades"
         Index           =   2
      End
      Begin VB.Menu M2101000000 
         Caption         =   "Registro de Compras"
         Index           =   3
      End
      Begin VB.Menu M2101000000 
         Caption         =   "Uso de Productos y Servicios"
         Index           =   4
      End
      Begin VB.Menu M2101000000 
         Caption         =   "Configurar Combos"
         Index           =   5
      End
      Begin VB.Menu M2101000000 
         Caption         =   "Entregas Directas"
         Index           =   6
      End
   End
End
Attribute VB_Name = "MDISicmact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'ALPA 20090122******************
Dim objPista As COMManejador.Pista
'*******************************
Private Sub M0101000000_Click(index As Integer)
    If index = 0 Then
        frmImpresora.Show 1
        'frmRHMantCtaCont.Show 1
    ElseIf index = 1 Then
        frmCaracImpresion.Show 1
    'ARLO20170511
    ElseIf index = 3 Then
                If MsgBox("Esta Seguro que Desea Salir del Sistema ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                Call SalirSICMACMAdmnstrativo
                End
            End If
        'Unload Me 'COMENTADO POR ARLO20170511
    End If
End Sub


Private Sub M0701010000_Click(index As Integer)
    'Persona
    Select Case index
        'Case 0 'Registro 'RECO20140312 ERS160-2013
            'frmPersona.Registrar
        'Case 1 'mantenimiento 'Registro 'RECO20140312 ERS160-2013
            'frmPersona.Mantenimeinto
        Case 2 'Consulta
            frmPersona.Consultar
    End Select
End Sub

Private Sub M0701020000_Click(index As Integer)
    
    'Instituciones Financieras
    Select Case index
        Case 0
            frmMntInstFinanc.InicioActualizar
        Case 1
            frmMntInstFinanc.InicioConsulta
    End Select
End Sub

Private Sub M0701030000_Click(index As Integer)
    If index = 0 Then
        frmPersEcoGruRel.Ini "PERSONAS:GRUPOS ECONOMICOS:MANTENIMIENTO"
    End If
End Sub

Private Sub M0701040000_Click(index As Integer)
    If index = 0 Then
        frmPromoRegistro.Show 1
    ElseIf index = 1 Then
        frmPromociones.Show 1
    End If
End Sub

Private Sub M0901000000_Click(index As Integer)
    If index = 0 Then
        frmEditorSicmact.Show 1
    ElseIf index = 1 Then 'Spooler
        frmSpooler.Show 1
    ElseIf index = 3 Then 'Perifericos
    
    ElseIf index = 5 Then 'BackUp
        
        frmBackUp.Show 1
    End If
    
End Sub

Private Sub M1001000000_Click(index As Integer)
    Select Case index
        Case 0
            frmMantPermisos.Show 1
        Case 1
            frmAdmUsu.Show 1
    End Select
End Sub

Private Sub M1601000000_Click(index As Integer)
    If index = 9 Then
        frmRHAsistenciaManual.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:ASISTENCIA:MANUAL"
    ElseIf index = 23 Then
        frmRHEmpleado.Ini gTipoOpeRegistro, RHContratoMantTpoComentario, "RECURSOS HUMANOS:COMENTARIO:REGISTRO"
    ElseIf index = 32 Then  '******* RECO 20130715******************
        frmRRHHRegistroExpedientes.Ini gTipoOpeRegistro, "RECURSOS HUMANOS:REGISTRO DE EXPEDIENTES:REGISTRO DE EXPEDIENTE DE RR.HH." '****RECO 20130715*****
    ElseIf index = 33 Then  '******* ARLO 20161015******************
        frmRHExpedienteDuración.Ini gTipoOpeRegistro, "REGISTRO EXPEDIENTE DURACIÓN "  '****ARLO 20161015*****
    End If
End Sub

Private Sub M1601010000_Click(index As Integer)
    If index = 3 Then
        frmRHSeleccion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuResultado, "RECURSOS HUMANOS:PROCESO SELECCION:RESULTADOS Y CIERRE"
    ElseIf index = 4 Then
        frmRHSeleccion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuResultado, "RECURSOS HUMANOS:PROCESO SELECCION:CONSULTA"
    End If
End Sub

Private Sub M1601010100_Click(index As Integer)
    If index = 0 Then
        frmRHSeleccion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuSel, "RECURSOS HUMANOS:PROCESO SELECCION:REGISTRO"
    ElseIf index = 1 Then
        frmRHSeleccion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuSel, "RECURSOS HUMANOS:PROCESO SELECCION:MANTENIMIENTO"
    ElseIf index = 2 Then
        frmRHSeleccion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuSel, "RECURSOS HUMANOS:PROCESO SELECCION:CONSULTA"
    ElseIf index = 3 Then
        frmRHSeleccion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuSel, "RECURSOS HUMANOS:PROCESO SELECCION:REPORTE"
    End If
End Sub

Private Sub M1601010200_Click(index As Integer)
    If index = 0 Then
        frmRHSeleccion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuPost, "RECURSOS HUMANOS:POSTULANTES:REGISTRO"
    ElseIf index = 1 Then
        frmRHSeleccion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuPost, "RECURSOS HUMANOS:POSTULANTES:MANTENIMIENTO"
    ElseIf index = 2 Then
        frmRHSeleccion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuPost, "RECURSOS HUMANOS:POSTULANTES:CONSULTA"
    ElseIf index = 3 Then
        frmRHSeleccion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuPost, "RECURSOS HUMANOS:POSTULANTES:REPORTE"
    End If
End Sub

Private Sub M1601010301_Click(index As Integer)
    If index = 0 Then
        frmRHSeleccion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaCur, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:CURRICULAR:REGISTRO"
    ElseIf index = 1 Then
        frmRHSeleccion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaCur, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:CURRICULAR:MANTENIMIENTO"
    ElseIf index = 2 Then
        frmRHSeleccion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaCur, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:CURRICULAR:CONSULTA"
    ElseIf index = 3 Then
        frmRHSeleccion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaCur, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:CURRICULAR:REPORTE"
    End If
End Sub

Private Sub M1601010302_Click(index As Integer)
    If index = 0 Then
        frmRHSeleccion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ESCRITO:REGISTRO"
    ElseIf index = 1 Then
        frmRHSeleccion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ESCRITO:MANTENIMIENTO"
    ElseIf index = 2 Then
        frmRHSeleccion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ESCRITO:CONSULTA"
    ElseIf index = 3 Then
        frmRHSeleccion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ESCRITO:REPORTE"
    End If
End Sub

Private Sub M1601010303_Click(index As Integer)
    If index = 0 Then
        frmRHSeleccion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaPsi, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:PSICOLOGICO:REGISTRO"
    ElseIf index = 1 Then
        frmRHSeleccion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaPsi, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:PSICOLOGICO:MANTENIMIENTO"
    ElseIf index = 2 Then
        frmRHSeleccion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaPsi, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:PSICOLOGICO:CONSULTA"
    ElseIf index = 3 Then
        frmRHSeleccion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaPsi, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:PSICOLOGICO:REPORTE"
    End If
End Sub

Private Sub M1601010304_Click(index As Integer)
    If index = 0 Then
        frmRHSeleccion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ENTREVISTA:REGISTRO"
    ElseIf index = 1 Then
        frmRHSeleccion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ENTREVISTA:MANTENIMIENTO"
    ElseIf index = 2 Then
        frmRHSeleccion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ENTREVISTA:CONSULTA"
    ElseIf index = 3 Then
        frmRHSeleccion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ENTREVISTA:REPORTE"
    End If
End Sub

Private Sub M1601020000_Click(index As Integer)
    If index = 0 Then
        frmRHContratoSeleccion.Ini "RECURSOS HUMANOS:CONTRATO:PROCESO SELECCION", ContratoFormaAutomatica
    ElseIf index = 1 Then
        frmRHContratoSeleccion.Ini "RECURSOS HUMANOS:CONTRATO:PROCESO SELECCION", ContratoFormaManual
    ElseIf index = 2 Then
        frmRHEmpleado.Ini gTipoOpeMantenimiento, RHContratoMantTpoFoto, "RECURSOS HUMANOS:CONTRATOS:MANTENIMIENTO"
    ElseIf index = 3 Then
        frmRHEmpleado.Ini gTipoOpeConsulta, RHContratoMantTpoFoto, "RECURSOS HUMANOS:FICHA PERSONAL:CONSULTA"
    ElseIf index = 4 Then
        Me.Enabled = False
        frmRHEmpleadoResCont.Ini "RECURSOS HUMANOS:CONTRATO:RESCINDIR CONTRATO", Me
        Me.Enabled = True
    End If
End Sub

Private Sub M1601030000_Click(index As Integer)
        If index = 0 Then
            frmRHEmpleado.Ini gTipoOpeRegistro, RHContratoMantTpoAdenda, "RECURSOS HUMANOS:ADENDA"
        ElseIf index = 1 Then
            frmFotocheck.Show
        End If
End Sub

Private Sub M1601040000_Click(index As Integer)
    If index = 1 Then
        frmRHCurriculum.Ini gTipoOpeRegistro, "RECURSOS HUMANOS:CURRICULUM VITAE:REGISTRO"
    ElseIf index = 2 Then
        frmRHCurriculum.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:CURRICULUM VITAE:MANTENIMIENTO"
    ElseIf index = 3 Then
        frmRHCurriculum.Ini gTipoOpeConsulta, "RECURSOS HUMANOS:CURRICULUM VITAE:CONSULTA"
    ElseIf index = 4 Then
        frmRHCurriculum.Ini gTipoOpeReporte, "RECURSOS HUMANOS:CURRICULUM VITAE:REPORTE"
    End If
End Sub

Private Sub M1601040100_Click(index As Integer)
    If index = 0 Then
        frmRHCurriculumTabla.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:CURRICULUM VITAE:TABLA:MANTENIMIENTO"
    ElseIf index = 1 Then
        frmRHCurriculumTabla.Ini gTipoOpeConsulta, "RECURSOS HUMANOS:CURRICULUM VITAE:TABLA:CONSULTA"
    End If
End Sub

Private Sub M1601050000_Click(index As Integer)
    If index = 0 Then
        frmRHAsistenciaAsig.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:HORARIO LABORAL:MANTENIMIENTO"
    ElseIf index = 1 Then
        frmRHAsistenciaAsig.Ini gTipoOpeConsulta, "RECURSOS HUMANOS:HORARIO LABORAL:CONSULTA"
    ElseIf index = 2 Then
        'frmRHAsistenciaHorario.Show
    End If
End Sub

Private Sub M1601060000_Click(index As Integer)
    If index = 5 Then
        frmRHEvaluacion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuResultado, "RECURSOS HUMANOS:EVA INT:RESULTADOS Y CIERRE"
    ElseIf index = 6 Then
        frmRHEvaluacion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuResultado, "RECURSOS HUMANOS:EVA INT:CONSULTA"
    End If
End Sub

Private Sub M1601060100_Click(index As Integer)
    If index = 0 Then
        frmRHEvaluacion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuSel, "RECURSOS HUMANOS:EVA INT:EVALUACION:Proceso EVALUACION Interna:REGISTRO"
    ElseIf index = 1 Then
        frmRHEvaluacion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuSel, "RECURSOS HUMANOS:EVA INT:EVALUACION:Proceso EVALUACION Interna:MANTENIMIENTO"
    ElseIf index = 2 Then
        frmRHEvaluacion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuSel, "RECURSOS HUMANOS:EVA INT:EVALUACION:Proceso EVALUACION Interna:CONSULTA"
    ElseIf index = 3 Then
        frmRHEvaluacion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuSel, "RECURSOS HUMANOS:EVA INT:EVALUACION:Proceso EVALUACION Interna:REPORTE"
    End If
End Sub

Private Sub M1601060200_Click(index As Integer)
    If index = 0 Then
        frmRHEvaluacion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaCur, "7RECURSOS HUMANOS:EVA INT:EVALUACION CURRICULAR:REGISTRO"
    ElseIf index = 1 Then
        frmRHEvaluacion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaCur, "RECURSOS HUMANOS:EVA INT:EVALUACION CURRICULAR:MANTENIMIENTO"
    ElseIf index = 2 Then
        frmRHEvaluacion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaCur, "RECURSOS HUMANOS:EVA INT:EVALUACION CURRICULAR:CONSULTA"
    ElseIf index = 3 Then
        frmRHEvaluacion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaCur, "RECURSOS HUMANOS:EVA INT:EVALUACION CURRICULAR:REPORTE"
    End If
End Sub

Private Sub M1601060300_Click(index As Integer)
    If index = 0 Then
        frmRHEvaluacion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:EVA INT:EVALUACION Escrita:REGISTRO"
    ElseIf index = 1 Then
        frmRHEvaluacion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:EVA INT:EVALUACION Escrita:MANTENIMIENTO"
    ElseIf index = 2 Then
        frmRHEvaluacion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:EVA INT:EVALUACION Escrita:CONSULTA"
    ElseIf index = 3 Then
        frmRHEvaluacion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:EVA INT:EVALUACION Escrita:REPORTE"
    End If
End Sub

Private Sub M1601060400_Click(index As Integer)
    If index = 0 Then
        frmRHEvaluacion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaPsi, "RECURSOS HUMANOS:EVA INT:EVALUACION Psicologica:REGISTRO"
    ElseIf index = 1 Then
        frmRHEvaluacion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaPsi, "RECURSOS HUMANOS:EVA INT:EVALUACION Psicologica:MANTENIMIENTO"
    ElseIf index = 2 Then
        frmRHEvaluacion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaPsi, "RECURSOS HUMANOS:EVA INT:EVALUACION Psicologica:CONSULTA"
    ElseIf index = 3 Then
        frmRHEvaluacion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaPsi, "RECURSOS HUMANOS:EVA INT:EVALUACION Psicologica:REPORTE"
    End If
End Sub

Private Sub M1601060500_Click(index As Integer)
    If index = 0 Then
        frmRHEvaluacion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:EVA INT:EVALUACION Entrevista:REGISTRO"
    ElseIf index = 1 Then
        frmRHEvaluacion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:EVA INT:EVALUACION Entrevista:MANTENIMIENTO"
    ElseIf index = 2 Then
        frmRHEvaluacion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:EVA INT:EVALUACION Entrevista:CONSULTA"
    ElseIf index = 3 Then
        frmRHEvaluacion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:EVA INT:EVALUACION Entrevista:REPORTE"
    End If
End Sub

Private Sub M1601070000_Click(index As Integer)
    If index = 1 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeMantenimiento, RHEstadosTpoPermisosLicencias, "RECURSOS HUMANOS:PERMISOS:APROBACION/RECHAZO"
    ElseIf index = 2 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeReporte, RHEstadosTpoPermisosLicencias, "RECURSOS HUMANOS:PERMISOS:APROBACION/RECHAZO"
    End If
End Sub

Private Sub M1601070100_Click(index As Integer)
    If index = 0 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeRegistro, RHEstadosTpoPermisosLicencias, ""
    ElseIf index = 1 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeConsulta, RHEstadosTpoPermisosLicencias, ""
    End If
End Sub

Private Sub M1601080000_Click(index As Integer)
    If index = 0 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeMantenimiento, RHEstadosTpoVacaciones, "RECURSOS HUMANOS:VACACIONES:MANTENIMIENTO"
    ElseIf index = 1 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeConsulta, RHEstadosTpoVacaciones, "RECURSOS HUMANOS:VACACIOBES:CONSULTA"
    End If
End Sub

Private Sub M1601090000_Click(index As Integer)
    If index = 0 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeMantenimiento, RHEstadosTpoSubsidiado, "RECURSOS HUMANOS:DESCANSOS:MANTENIMIENTO"
    ElseIf index = 1 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeConsulta, RHEstadosTpoSubsidiado, "RECURSOS HUMANOS:DESCANSOS:CONSULTA"
    ElseIf index = 2 Then
        frmRHReportesSubsidio.Show 1
    End If
End Sub

Private Sub M1601100000_Click(index As Integer)
    If index = 0 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeMantenimiento, RHEstadosTpoSuspendido, "RECURSOS HUMANOS:SANCIONES:MANTENIMIENTO"
    ElseIf index = 1 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeConsulta, RHEstadosTpoSuspendido, "RECURSOS HUMANOS:SANCIONES:CONSULTA"
    End If
End Sub

Private Sub M1601110000_Click(index As Integer)
    If index = 1 Then
        frmRHMerDem.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:MERITOS Y DEMERITOS:TABLA:MANTENIMIENTO"
    ElseIf index = 2 Then
        frmRHMerDem.Ini gTipoOpeConsulta, "RECURSOS HUMANOS:MERITOS Y DEMERITOS:TABLA:CONSULTA"
    End If
End Sub

Private Sub M1601110100_Click(index As Integer)
    If index = 0 Then
        frmRHMerDemTabla.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:MERITOS Y DEMERITOS:TABLA:MANTENIMIENTO"
    ElseIf index = 1 Then
        frmRHMerDemTabla.Ini gTipoOpeReporte, "RECURSOS HUMANOS:MERITOS Y DEMERITOS:TABLA:CONSULTA"
    End If
End Sub

Private Sub M1601120000_Click(index As Integer)
    If index = 1 Then
        frmRHEmpleado.Ini gTipoOpeRegistro, RHContratoMantTpoCargo, "RECURSOS HUMANOS:CARGOS LABORALES:REGISTRO"
    End If
End Sub

Private Sub M1601120100_Click(index As Integer)
    If index = 0 Then
        frmRHCargos.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:CARGOS LABORALES:TABLA:MANTENIMIENTO"
    ElseIf index = 1 Then
        frmRHCargos.Ini gTipoOpeConsulta, "RECURSOS HUMANOS:CARGOS LABORALES:TABLA:CONSULTA"
    End If
End Sub

Private Sub M1601130000_Click(index As Integer)
    If index = 0 Then
        frmRHEmpleado.Ini gTipoOpeRegistro, RHContratoMantTpoSueldo, "RECURSOS HUMANOS:SUELDO:REGISTRO"
    End If
End Sub

Private Sub M1601140000_Click(index As Integer)
    If index = 1 Then
        frmRHEmpleado.Ini gTipoOpeRegistro, RHContratoMantTpoSisPens, "RECURSOS HUMANOS:SISTEMA PENSIONES:REGISTRO"
    End If
End Sub

Private Sub M1601140100_Click(index As Integer)
    If index = 0 Then
        frmRHAFP.Ini "RECURSOS HUMANOS:SISTEMA PENSIONES:TABLA:MANTENIMIENTO"
    End If
End Sub

Private Sub M1601150000_Click(index As Integer)
    If index = 0 Then
        frmRHInformeSocial.Ini gTipoOpeRegistro, "RECURSOS HUMANOS:INFORME SOCIAL:REGISTRO"
    ElseIf index = 1 Then
        frmRHInformeSocial.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:INFORME SOCIAL:MANTENIMIENTO"
    ElseIf index = 2 Then
        frmRHInformeSocial.Ini gTipoOpeConsulta, "RECURSOS HUMANOS:INFORME SOCIAL:CONSULTA"
    ElseIf index = 3 Then
        frmRHInformeSocial.Ini gTipoOpeReporte, "RECURSOS HUMANOS:INFORME SOCIAL:REPORTE"
    End If
End Sub

Private Sub M1601160000_Click(index As Integer)
    If index = 1 Then
        frmRHEmpleado.Ini gTipoOpeRegistro, RHContratoMantTpoAMP, "RECURSOS HUMANOS:ASISTENCIA MEDICA;REGISTRO"
        'frmRHAsignacionPlan.Show
    End If
End Sub

Private Sub M1601160100_Click(index As Integer)
    If index = 0 Then
        frmRHAsistMedPrivada.Ini "RECURSOS HUMANOS:ASITENCIA MEDICA:TABLA:MANTENIMIENTO"
    End If
End Sub

Private Sub M1601170100_Click(index As Integer)
    If index = 0 Then
        frmRHConceptoMant.Ini gTipoOpeMantenimiento, "RRHH:CONFIG. DE CONCEPTOS REMUNERATIVOS:MANTENIMIENTO"
    ElseIf index = 1 Then
        frmRHTablasAlias.Ini "RRHH:CONFIG. DE CONCEPTOS REMUNERATIVOS:MANTENIMIENTO TABLA ALIAS"
    ElseIf index = 2 Then
        frmRHConceptoMant.Ini gTipoOpeConsulta, "RRHH:CONFIG. DE CONCEPTOS REMUNERATIVOS:CONSULTA"
    ElseIf index = 3 Then
        frmRHConceptoMant.Ini gTipoOpeReporte, "RRHH:CONFIG. DE CONCEPTOS REMUNERATIVOS:REPORTE"
    ElseIf index = 4 Then
        frmRHMantCtaCont.Ini "RRHH:CONFIG. DE CONCEPTOS REMUNERATIVOS:MANT.CTAS.CONTABLES"
    End If
End Sub

Private Sub M1601170200_Click(index As Integer)
    If index = 0 Then
        frmRHConceptoAsigPla.Ini gTipoOpeRegistro, "RRHH:CONFIG. PLANILLA DE REMUNERACIONES:REGISTRO"
    ElseIf index = 1 Then
        frmRHConceptoAsigPla.Ini gTipoOpeMantenimiento, "RRHH:CONFIG. PLANILLA DE REMUNERACIONES:MANTENIMIENTO"
    ElseIf index = 2 Then
        frmRHConceptoAsigPla.Ini gTipoOpeConsulta, "RRHH:CONFIG. PLANILLA DE REMUNERACIONES:CONSULTA"
    ElseIf index = 3 Then
        frmRHConceptoAsigPla.Ini gTipoOpeReporte, "RRHH:CONFIG. PLANILLA DE REMUNERACIONES:REPORTE"
    End If
End Sub

Private Sub M1601180100_Click(index As Integer)
    If index = 0 Then
        frmRHConceptoAsigRRHH.Ini gTipoOpeMantenimiento, True, "RRHH:PLANILLA DE REMM RRHH:CONCEPTOS FIJOS:MANTENIMIENTO"
    ElseIf index = 1 Then
        frmRHConceptoAsigRRHH.Ini gTipoOpeConsulta, True, "RRHH:PLANILLA DE REMM RRHH:CONCEPTOS FIJOS:CONSULTA"
    ElseIf index = 2 Then
        frmRHConceptoAsigRRHH.Ini gTipoOpeReporte, True, "RRHH:PLANILLA DE REMM RRHH:CONCEPTOS FIJOS:REPORTE"
    ElseIf index = 3 Then
        frmRHPrePlanilla.Show
    End If
End Sub

Private Sub M1601180200_Click(index As Integer)
    If index = 0 Then
        frmRHConceptoAsigRRHH.Ini gTipoOpeMantenimiento, False, "RRHH:PLANILLA DE REMM RRHH:CONCEPTOS VARIABLES:MANTENIMIENTO"
    ElseIf index = 1 Then
        frmRHConceptoAsigRRHH.Ini gTipoOpeConsulta, False, "RRHH:PLANILLA DE REMM RRHH:CONCEPTOS VARIABLES:CONSULTA"
    ElseIf index = 2 Then
        frmRHConceptoAsigRRHH.Ini gTipoOpeReporte, False, "RRHH:PLANILLA DE REMM RRHH:CONCEPTOS VARIABLES:REPORTE"
    End If
End Sub

Private Sub M1601180300_Click(index As Integer)
    If index = 0 Then
        frmRHPlanillaMovExtra.Ini gTipoOpeRegistro, "RECURSOS HUMANOS:PLANILLA DE REMUNERACIONES:EXTRA PLANILLA:REGISTRO"
    ElseIf index = 1 Then
        frmRHPlanillaMovExtra.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:PLANILLA DE REMUNERACIONES:EXTRA PLANILLA:MANTENIMIENTO"
    ElseIf index = 2 Then
        frmRHPlanillaMovExtra.Ini gTipoOpeConsulta, "RECURSOS HUMANOS:PLANILLA DE REMUNERACIONES:EXTRA PLANILLA:CONSULTA"
    ElseIf index = 3 Then
        frmRHPlanillaMovExtra.Ini gTipoOpeReporte, "RECURSOS HUMANOS:PLANILLA DE REMUNERACIONES:EXTRA PLANILLA:REPORTE"
    End If
End Sub

Private Sub M1601180400_Click(index As Integer)
    If index = 0 Then
        frmRHPrestamosAdm.Show 1
    ElseIf index = 1 Then
        frmRHPrestamosAdm.Show 1
    End If
End Sub

Private Sub M1601180500_Click(index As Integer)
    If index = 0 Then
        frmRHPrestamosADMOtros.Show 1
    ElseIf index = 1 Then
        frmRHPrestamosADMOtros.Show 1
    End If
End Sub

Private Sub M1601190000_Click(index As Integer)
    Me.Enabled = False
    If index = 0 Then
        frmRHPlanillas.Ini gTipoProcesoRRHHCalculo, "RECURSOS HUMANOS:PROCESOS:CALCULO DE PLANILLAS", Me
    ElseIf index = 1 Then
        frmRHPlanillas.Ini gTipoProcesoRRHHAbono, "RECURSOS HUMANOS:PROCESOS:ABONO DE PLANILLAS", Me
    ElseIf index = 2 Then
        frmRHCierreMes.Ini "RECURSOS HUMANOS:PROCESOS:CIERRE MES"
    ElseIf index = 3 Then
        frmRHCierreDia.Ini "RECURSOS HUMANOS:PROCESOS:CIERRE DIA"
    ElseIf index = 4 Then
        frmRHPagoAMP.Show 1
    ElseIf index = 5 Then
        frmRHPeriodosNoLaboralesProc.Show 1
    End If
    Me.Enabled = True
End Sub



Private Sub M1601190002_Click(index As Integer)
Select Case index
    Case 0
        frmVacacionesGozadas.Show 1
        
    Case 1
        frmProvisionCTS.Show 1
        
    Case 2
        frmProvisionGratificacion.Show 1
End Select


End Sub

Private Sub M1601200000_Click(index As Integer)
  
    
  

    Me.Enabled = False
    If index = 0 Then
        frmRRHHRep.Ini "RECURSOS HUMANOS:REPORTES:REPORTES", Me
    ElseIf index = 1 Then
        frmRRHHRepGen.Ini "RECURSOS HUMANOS:REPORTES:REPORTES GENERALES", Me
    ElseIf index = 2 Then        '************RECO 20130727************************
        frmRRHHRepMovPersonal.Ini "RECURSOS HUMANOS:REPORTE:REPORTE MOVIMIENTO DE PERSONAL", Me
    ElseIf index = 3 Then
        MsgBox "Opción se movio al sistema de reportes", vbInformation, "aviso" 'PTI1 ERS029 26042018
        'frmRRHHRepExpedientes.Ini "RECURSOS HUMANOS:REPORTE:REPORTE DE EXPEDIENTE", Me 'PTI1 ERS029 26042018
        '****END RECO********************************
    End If
    Me.Enabled = True

    
End Sub

'Private Sub M0101000000_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            frmImpresora.Show 1
'        Case 1
'            If MsgBox("Esta Seguro que Desea Salir del Sistema ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
'                End
'            End If
'    End Select
'
'End Sub
'
'Private Sub M0201000000_Click(Index As Integer)
'    Select Case Index
'    Case 7 'Refinanciacion de Credito
'        Call frmCredSolicitud.RefinanciaCredito(Registrar)
'    Case 8 ' 'Actualizacion con Metodos de Liquidacion
'        frmCredMntMetLiquid.Show 1
'    End Select
'End Sub
'
'Private Sub M0201010100_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Mantenimeinto de Parametros
'            frmCredMantParametros.InicioActualizar
'        Case 1 'Consulta de Parametros
'            frmCredMantParametros.InicioCosultar
'    End Select
'End Sub
'
'Private Sub M0201010200_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Registro de Lineas de Credito
'            frmCredLineaCredito.Registrar
'        Case 1 'Mantenimiento de Lineas de Credito
'            frmCredLineaCredito.Actualizar
'        Case 2 ' Consulta de lineas de Credito
'            frmCredLineaCredito.Consultar
'    End Select
'End Sub
'
'Private Sub M0201010300_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Mantenimeinto de Niveles de Aprobacion
'            frmCredNivAprCred.inicio MuestraNivelesActualizar
'        Case 1 'Consulta de Niveles de Aprobacion
'            frmCredNivAprCred.inicio MuestraNivelesConsulta
'    End Select
'End Sub
'
'Private Sub M0201010400_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Mantenimeinto de Gastos
'            frmCredMntGastos.inicio InicioGastosActualizar
'        Case 1
'            frmCredMntGastos.inicio InicioGastosConsultar
'    End Select
'End Sub
'
'Private Sub M0201020000_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Registro de Solicitud
'            frmCredSolicitud.inicio Registrar
'        Case 1 'Consulta de Solicitud
'            frmCredSolicitud.inicio Consulta
'    End Select
'End Sub
'
'Private Sub M0201030000_Click(Index As Integer)
'Dim oCredRel As New UCredRelacion
'
'    Select Case Index
'        Case 0 'Mantenimiento de Relaciones de Credito
'            frmCredRelaCta.inicio oCredRel, InicioMantenimiento, InicioRegistroForm
'            Set oCredRel = Nothing
'        Case 1 'Reasignacion de Cartera en Lote
'            frmCredReasigCartera.Show 1
'        Case 2 'Consulta de Relaciones de Credito
'            frmCredRelaCta.inicio oCredRel, InicioMantenimiento, InicioConsultaForm
'            Set oCredRel = Nothing
'    End Select
'End Sub
'
'Private Sub M0201040000_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Registro de Garantia
'            frmPersGarantias.inicio RegistroGarantia
'        Case 1 'Mantenimiento de Garantia
'            frmPersGarantias.inicio MantenimientoGarantia
'        Case 2 'Consulta de Garantia
'            frmPersGarantias.inicio ConsultaGarant
'        Case 3 'Gravament
'            frmCredGarantCred.inicio PorMenu
'    End Select
'End Sub
'
'Private Sub M0201050000_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Registro de Sugerencia
'            frmCredSugerencia.Show 1
'        Case 1 'Mantenimiento de Sugerencia
'
'        Case 2 'Consulta de Sugerencia
'
'    End Select
'End Sub
'
'Private Sub M0201060000_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Aprobacion de Credito
'            frmCredAprobacion.Show 1
'        Case 1 'Rechazo de Credito
'            frmCredRechazo.Rechazar
'        Case 2 'Anulacion de Credito
'            frmCredRechazo.Retirar
'    End Select
'End Sub
'
'Private Sub M0201070000_Click(Index As Integer)
'    Select Case Index
'        Case 0 'Reprogramacion de Credito
'            frmCredReprogCred.Show 1
'        Case 1 'Reprogramacion en Lote
'            frmCredReprogLote.Show 1
'    End Select
'End Sub
'
'
'
'
'
'Private Sub M1701000000_Click(Index As Integer)
'
'End Sub
'
'Private Sub MDIForm_Unload(Cancel As Integer)
'    If MsgBox("Esta Seguro que Desea Salir del Sistema ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
'        Cancel = 1
'    End If
'End Sub
'
'Private Sub mnuCapBenCons04070100_Click()
'    frmCapBeneficiario.Inicia True
'End Sub
'
'Private Sub mnuCapBenMnt04070200_Click()
'    frmCapBeneficiario.Inicia False
'End Sub
'
'Private Sub mnuCapBloqAho04040100_Click()
'    frmCapBloqueoDesbloqueo.Inicia gCapAhorros
'End Sub
'
'Private Sub mnuCapBloqCTS04040300_Click()
'    frmCapBloqueoDesbloqueo.Inicia gCapCTS
'End Sub
'
'Private Sub mnuCapBloqPF04040200_Click()
'    frmCapBloqueoDesbloqueo.Inicia gCapPlazoFijo
'End Sub
'
'Private Sub mnuCapMntAho04030100_Click()
'    frmCapMantenimiento.Inicia gCapAhorros
'End Sub
'
'Private Sub mnuCapMntCTS04030300_Click()
'    frmCapMantenimiento.Inicia gCapCTS
'End Sub
'
'Private Sub mnuCapMntPF04030200_Click()
'    frmCapMantenimiento.Inicia gCapPlazoFijo
'End Sub
'
'Private Sub mnuCapOrdAnu04080300_Click()
'    frmCapOrdPagAnulCert.Inicia gAhoOPAnulacion
'End Sub
'
'Private Sub mnuCapOrdCer04080200_Click()
'    frmCapOrdPagAnulCert.Inicia gAhoOPCertificacion
'End Sub
'
'Private Sub mnuCapOrdCon04080400_Click()
'    frmCapOrdPagConsulta.Show 1
'End Sub
'
'Private Sub mnuCapOrdGen04080100_Click()
'    frmCapOrdPagGenEmi.Show 1
'End Sub
'
'Private Sub mnuCapParCons04010100_Click()
'    frmCapParametros.Inicia True
'End Sub
'
'Private Sub mnuCapParMnt04010200_Click()
'    frmCapParametros.Inicia False
'End Sub
'
'Private Sub mnuCapSerConvCta04090402_Click()
'    frmCapServConvCuentas.Inicia
'End Sub
'
'Private Sub mnuCapSerConvMnt04090401_Click()
'    frmCapConvenioMant.Show 1
'End Sub
'
'Private Sub mnuCapSerConvPln04090403_Click()
'    frmCapServConvPlanPag.Inicia
'End Sub
'
'Private Sub mnuCapSerHidPar04090201_Click()
'    frmCapServParametros.Inicia gCapServHidrandina
'End Sub
'
'Private Sub mnuCapSerHidRep04090202_Click()
'    frmCapServGenReporte.Inicia gCapServHidrandina
'End Sub
'
'Private Sub mnuCapSerSedGen04090103_Click()
'    frmCapServGeneraDBF.Show 1
'End Sub
'
'Private Sub mnuCapSerSedPar04090101_Click()
'    frmCapServParametros.Inicia gCapServSedalib
'End Sub
'
'Private Sub mnuCapSerSedRep04090102_Click()
'    frmCapServGenReporte.Inicia gCapServSedalib
'End Sub
'
'Private Sub mnuCapSimPF04050000_Click()
'    frmCapSimulacionPF.Show 1
'End Sub
'
'Private Sub mnuCapTarBloq04060300_Click()
'    frmCapTarjetaBlqDesBlq.Show 1
'End Sub
'
'Private Sub mnuCapTarCla04060400_Click()
'    frmCapTarjetaCambioClave.Show 1
'End Sub
'
'Private Sub mnuCapTarReg04060100_Click()
'    frmCapTarjetaRegistro.Inicia
'End Sub
'
'Private Sub mnuCapTarRel04060200_Click()
'    frmCapTarjetaRelacion.Inicia False
'End Sub
'
'Private Sub mnuCapTasaCTSCons04020301_Click()
'    frmCapTasaInt.Inicia gCapCTS, True
'End Sub
'
'Private Sub mnuCapTasaCTSCons04020302_Click()
'    frmCapTasaInt.Inicia gCapCTS, False
'End Sub
'
'Private Sub mnuCapTasAhoCons04020101_Click()
'    frmCapTasaInt.Inicia gCapAhorros, True
'End Sub
'
'Private Sub mnuCapTasAhoMnt04020102_Click()
'    frmCapTasaInt.Inicia gCapAhorros, False
'End Sub
'
'Private Sub mnuCapTasaPFCons04020201_Click()
'    frmCapTasaInt.Inicia gCapPlazoFijo, True
'End Sub
'
'Private Sub mnuCapTasaPFCons04020202_Click()
'    frmCapTasaInt.Inicia gCapPlazoFijo, False
'End Sub
'
'Private Sub mnuCliPosic08010000_Click()
'    frmPosicionCli.Show 1
'End Sub
'
'Private Sub mnucredAnaMetas02150200_Click()
'    frmCredMetasAnalista.Show 1
'End Sub
'
'Private Sub mnucredAnaNota02150100_Click()
'    frmCredAsigNota.Show 1
'End Sub
'
'Private Sub mnucredGastosLote02110100_Click()
'    frmCredAsigGastosLote.Show 1
'End Sub
'
'Private Sub mnucredGastosPenalidad02110200_Click()
'    frmCredExonerarPen.Show 1
'End Sub
'
'Private Sub mnucredHistorial02160100_Click()
'    frmCredConsulta.Show 1
'End Sub
'
'Private Sub mnucredPasarAJud02140000_Click()
'    frmCredTransARecup.Show 1
'End Sub
'
'Private Sub mnucredPerdMora02100000_Click()
'    frmCredPerdonarMora.Show 1
'End Sub
'
'
'
'Private Sub mnucredReasignarInst02130000_Click()
'    frmCredReasigInst.Show 1
'End Sub
'
'
'
'Private Sub mnucredRepDupDocum02170100_Click()
'    frmCredDupDoc.Show 1
'End Sub
'
'
'Private Sub mnucredSimCalCuoLib02090300_Click()
'Dim MatCalend As Variant
'Dim Matriz(0) As String
'    MatCalend = frmCredCalendCuotaLibre.CalendarioLibre(True, gdFecSis, Matriz, 0#, 0, 0#)
'End Sub
'
'Private Sub mnucredSimCalDesPar02090200_Click()
'    frmCredCalendPagos.Simulacion DesembolsoParcial
'End Sub
'
'Private Sub mnucredSimCalPag02090100_Click()
'    frmCredCalendPagos.Simulacion DesembolsoTotal
'End Sub
'
'
'
'Private Sub mnuHerramPerif09020000_Click()
'    frmSetupCOM.Show 1
'End Sub
'
'Private Sub mnuHerramSpooler09010000_Click()
'    frmSpooler.Show 1
'End Sub
'
'Private Sub mnuIFinanCons07030200_Click()
'    frmMntInstFinanc.InicioConsulta
'End Sub
'
'
'Private Sub mnuOpCajero06900000_Click()
'    frmCajeroOperaciones.Show 1
'End Sub
'
'Private Sub mnuOpCajeroCMACLLam06920000_Click()
'    frmCajeroOpeCMAC.Inicia False
'End Sub
'
'Private Sub mnuOpCajeroCMACRes06910000_Click()
'    frmCajeroOpeCMAC.Inicia
'End Sub
'
'Private Sub mnuOpCajeroExtCapta06930400_Click()
'    frmCapExtornos.Show 1
'End Sub
'
'Private Sub mnuOpeDesemAbo06010000_Click()
'    frmCredDesembAbonoCta.DesembolsoCargoCuenta
'End Sub
'
'Private Sub mnupersonaCons07010300_Click()
'    frmPersona.Consultar
'End Sub
'
'Private Sub mnupersonamant07010200_Click()
'    frmPersona.Mantenimeinto
'End Sub
'
'Private Sub mnupersonareg07010100_Click()
'    frmPersona.Registrar
'End Sub
'
'Private Sub mnuPigAdjudica03040000_Click()
'    frmColPAdjudicaLotes.Show 1
'End Sub
'
'Private Sub mnuPigAnulacion03010300_Click()
'    frmColPAnularPrestamoPig.Show
'End Sub
'
'Private Sub mnuPigBloqueo03010400_Click()
'    frmColPBloqueo.Show 1
'End Sub
'
'Private Sub mnuPigContraRegis03010100_Click()
'    frmColPRegContrato.Show 1
'End Sub
'
'Private Sub mnuPigMntgDescrip03010200_Click()
'    frmColPMantPrestamoPig.Show
'End Sub
'
'Private Sub mnuPigRemPrepRem03030100_Click()
'     frmColPRematePrepara.Show 1
'End Sub
'
'Private Sub mnuPigRemRem03030200_Click()
'    frmColPRemateProceso.Show 1
'End Sub
'
'Private Sub mnuPigRescate03020000_Click()
'    frmColPRescateJoyas.Show 1
'End Sub
'
'Private Sub mnuPigSubPrep03050100_Click()
'    frmColPSubastaPrepara.Show 1
'End Sub
'
'Private Sub mnuPigSubSubasta03050200_Click()
'    frmColPSubastaProceso.Show 1
'End Sub
'
'Private Sub mnuSegurPerm10010000_Click()
'    frmMantPermisos.Show 1
'End Sub
'
'Private Sub mnutFinanMant07030100_Click()
'    frmMntInstFinanc.InicioActualizar
'End Sub
'
'Private Sub Tiempo_Timer()
'    staMain.Panels(2).Text = Format(gdFecSis, "dddd - dd - mmmm - yyyy") & Space(3) & Format(Time, "hh:mm AMPM")
'End Sub

Private Sub M1701000000_Click(index As Integer)
    If index = 0 Then
        'frmLogUsuario.Show 1
        'frmLogBSActFijoMant.Show 1
        'frmLogAfTrans.Show 1
    ElseIf index = 2 Then
        frmLogBieSerMant_des.Show 1
    ElseIf index = 4 Then
        'frmPasePersona.Show
        frmLogProvMant.Show 1
    ElseIf index = 13 Then
        'frmLogSelConsol.Inicio "3"
        frmLogSelEvalTecResumen.Inicio "R", "1"
    ElseIf index = 14 Then
        frmLogSelCancelacionProceso.Show
    ElseIf index = 15 Then
        'frmLogSelConsol.Inicio "4"
    ElseIf index = 16 Then
        'frmLogSelConsol.Inicio "5"
    ElseIf index = 17 Then
        'frmLogSelConsol.Inicio "1"
    'ElseIf Index = 22 Then
    '    frmLogAFDeprecia.Ini "RECURSOS HUMANOS:DEPRECIACION DE ACTIVO FIJO"
    ElseIf index = 24 Then
        frmLogOperacionesIngBS.Inicio "58"
    'ElseIf Index = 24 Then
    '    frmTransferencia.Show 1
    'ElseIf Index = 25 Then
    '    frmAFAsiganPersona.Show 1
    'ElseIf Index = 26 Then
    '    frmAFBajaActivo.Show 1
    ElseIf index = 27 Then
        frmLogRetasacion.Show 1
    ElseIf index = 28 Then
        frmLogSaneamiento.Show 1
    ElseIf index = 30 Then
        frmLogMantBienesAdjud.Show 1
    End If
End Sub

Private Sub M1701010000_Click(index As Integer)
    If index = 0 Then
        frmLogReqInicio.Inicio "1", "1"
    ElseIf index = 1 Then
        frmLogReqTramite.Inicio "1"
    ElseIf index = 2 Then
        'frmLogReqPrecio.Inicio "1", "1"
    ElseIf index = 3 Then
        'frmLogReqPrecio.Inicio "1", "3"
    ElseIf index = 4 Then
        'frmLogReqPrecio.Inicio "1", "2"
    End If
End Sub

Private Sub M1701020000_Click(index As Integer)
    If index = 0 Then
        frmLogReqInicio.Inicio "2", "1"
    ElseIf index = 1 Then
        frmLogReqTramite.Inicio "2"
    ElseIf index = 2 Then
        'frmLogReqPrecio.Inicio "2", "1"
    ElseIf index = 3 Then
        'frmLogReqPrecio.Inicio "2", "2"
    ElseIf index = 4 Then
        'frmLogReqPrecio.Inicio "2", "3"
    End If
End Sub

Private Sub M1701030000_Click(index As Integer)
    
   If index = 0 Then
    frmLogReqConsolidacion.Show
    ElseIf index = 1 Then
    frmLogReqAprobacion.Show
    End If
End Sub


Private Sub M1701040000_Click(index As Integer)
    Select Case index
    Case 0 'Comite
         frmLogSelComite.Show
    'Case 1
         'frmLogSelComienzo.Show
    Case 2
         frmLogSelProveedores.Show
    Case 3
         frmLogSelSeleccionBienes.Inicio "1", "1"
    End Select
End Sub

Private Sub M1701040101_Click(index As Integer)
    'frmLogSelInicio.Inicio IIf(Index = 0, "1", "2"), IIf(Index = 0, False, True)
End Sub

Private Sub M1701040102_Click(index As Integer)
    'frmLogSelInicio.Inicio IIf(Index = 0, "2", "3"), IIf(Index = 0, False, True)
End Sub

Private Sub M1701040103_Click(index As Integer)
    'frmLogSelInicio.Inicio IIf(Index = 0, "3", "4"), IIf(Index = 0, False, True)
End Sub

Private Sub M1701040104_Click(index As Integer)
    'frmLogSelInicio.Inicio IIf(Index = 0, "4", "5"), IIf(Index = 0, False, True)
End Sub

Private Sub M1701040105_Click(index As Integer)
    'frmLogSelInicio.Inicio IIf(Index = 0, "5", "6"), IIf(Index = 0, False, True)
End Sub

Private Sub M1701040106_Click(index As Integer)
    'frmLogSelInicio.Inicio IIf(Index = 0, "6", "7"), IIf(Index = 0, False, True)
End Sub

Private Sub M1701040301_Click(index As Integer)
    'frmLogSelInicio.Inicio IIf(Index = 0, "7", "8"), IIf(Index = 0, False, True)
End Sub

Private Sub M1701040302_Click(index As Integer)
    'frmLogSelInicio.Inicio IIf(Index = 0, "8", "9"), IIf(Index = 0, False, True)
End Sub

Private Sub M1701040303_Click(index As Integer)
    'frmLogSelInicio.Inicio IIf(Index = 0, "9", "10"), IIf(Index = 0, False, True)
End Sub

Private Sub M1701040304_Click(index As Integer)
    'frmLogSelInicio.Inicio IIf(Index = 0, "10", "11"), IIf(Index = 0, False, True)
End Sub

Private Sub M1701040300_Click(index As Integer)
If index = 0 Then
    frmLogSelCriterios.Show
ElseIf index = 1 Then
    frmLogSelAsignacionCriterios.Show
End If
End Sub

Private Sub M1701040400_Click(index As Integer)
    If index = 0 Then
        frmLogSelSeleccionBienes.Inicio "2", "2"
    ElseIf index = 1 Then
        'frmLogSelSeleccionBienes.Inicio "3", "3"
    ElseIf index = 2 Then
        'frmLogSelCotPro.inicio "1", "3"
    End If
End Sub
Private Sub M1701040650_Click(index As Integer)
If index = 0 Then
        'Evaluacion Tecnica Registro
        frmLogEvaluacionTecnica.Show
ElseIf index = 1 Then
        'Evaluacion Tecnica Resumen
        frmLogSelEvalTecResumen.Inicio "T", "1"
End If
End Sub

Private Sub M1701040660_Click(index As Integer)
If index = 0 Then
frmLogSelSeleccionBienes.Inicio "3", "3"
ElseIf index = 1 Then
frmLogSelEvalTecResumen.Inicio "E", "1"
End If
End Sub

'TORE RFC1902190004 - Carta Fianza
Private Sub M1701050000_Click(index As Integer)
    If index = 8 Then
        frmCartaFianza.Show 1
    End If
End Sub

Private Sub M1701050100_Click(index As Integer)
    If index = 0 Then
        frmLogOCompra.Inicio False, "501205"
    ElseIf index = 1 Then
        frmLogOCompra.Inicio False, "502205"
    ElseIf index = 2 Then
        frmLogOCompra.Inicio False, "501207", "", False
    ElseIf index = 3 Then
        frmLogOCompra.Inicio False, "502207", "", False
    End If
    Set frmLogOCompra = Nothing
End Sub

Private Sub M1701050200_Click(index As Integer)
    If index = 0 Then
        frmLogOCAtencion.Inicio True, "501206", False, True
    ElseIf index = 1 Then
        frmLogOCAtencion.Inicio True, "502206", False, True
    ElseIf index = 2 Then
        frmLogOCAtencion.Inicio False, "501208", False, True
    ElseIf index = 3 Then
        frmLogOCAtencion.Inicio False, "502208", False, True
    ElseIf index = 4 Then 'PASI20151231
        frmLogOrdenesSeguimiento.Show 1
    End If
    Set frmLogOCAtencion = Nothing
End Sub

Private Sub M1701050300_Click(index As Integer)
    If index = 0 Then
        frmLogOCAtencion.Inicio True, "501210", False, False, True
    ElseIf index = 1 Then
        frmLogOCAtencion.Inicio True, "502210", False, False, True
    ElseIf index = 2 Then
        frmLogOCAtencion.Inicio False, "501211", False, False, True
    ElseIf index = 3 Then
        frmLogOCAtencion.Inicio False, "502211", False, False, True
    End If
    Set frmLogOCAtencion = Nothing
End Sub

Private Sub M1701050400_Click(index As Integer)
    If index = 0 Then
        frmLogOCAtencion.Inicio True, "501210", False, False, True, True
    ElseIf index = 1 Then
        frmLogOCAtencion.Inicio True, "502210", False, False, True, True
    ElseIf index = 2 Then
        frmLogOCAtencion.Inicio False, "501211", False, False, True, True
    ElseIf index = 3 Then
        frmLogOCAtencion.Inicio False, "502211", False, False, True, True
    End If
    Set frmLogOCAtencion = Nothing
End Sub

Private Sub M1701060000_Click(index As Integer)
    If index = 0 Then
        frmLogOCAtencion.Inicio True, "501221"
    ElseIf index = 1 Then
        frmLogOCAtencion.Inicio False, "501222"
    ElseIf index = 2 Then
        frmLogOCAtencion.Inicio True, "502221"
    ElseIf index = 3 Then
        frmLogOCAtencion.Inicio False, "502222"
    End If
    Set frmLogOCAtencion = Nothing
End Sub
'WIOR 20120806 *********************************
Private Sub M1701050500_Click(index As Integer)
Select Case index
    'REGISTRO DE CONTRATOS
    Case 0: frmLogContRegistro.Show 1
    'SEGUIMIENTO DE CONTRATOS
    Case 1:  frmLogContSeguimiento.Show 1
End Select
End Sub

'WIOR FIN **************************************
'EJVG20131105 *** Se deshabilito anterior Registro de Contrato
'WIOR 20130110 *********************************
'Private Sub M1701050600_Click(index As Integer)
'Select Case index
'    Case 0: frmLogOCAtencion.Inicio True, "501210", False, False, False, , True
'    Case 1: frmLogOCAtencion.Inicio True, "502210", False, False, False, , True
'    Case 2: frmLogOCAtencion.Inicio False, "501211", False, False, False, , True
'    Case 3: frmLogOCAtencion.Inicio False, "502211", False, False, False, , True
'    Case 4: frmLogContRegComprobantes.Inicio
'    Case 5: frmLogContAndRegComp.Inicio
'    Case 6: frmLogOCompra.Inicio False, "501205" ', False, False, False, , True
'    Case 7: frmLogImpComprobantes.Inicio
'End Select
'Set frmLogOCAtencion = Nothing
'End Sub
Private Sub M1701050700_Click(index As Integer)
       'Dim oActa As frmLogActaConformidad 'Comentado PASI20140925 ERS0772014
    Dim oActa As frmLogActaConformidadNew
    'Dim oLibre As frmLogActaConformidadLibre 'Comentado PASI20140925 ERS0772014
    Dim oExtorno As frmLogActaConformidadExtorno
    Select Case index
        Case 0:
            'Set oActa = New frmLogActaConformidad 'Comentado PASI20140925 ERS0772014
            Set oActa = New frmLogActaConformidadNew 'PASI20140925 ERS0772014
            oActa.Inicio gnAlmaActaConformidadMN, "ACTA DE CONFORMIDAD DIGITAL SOLES"
        Case 1:
            'Set oActa = New frmLogActaConformidad 'Comentado PASI20140925 ERS0772014
            Set oActa = New frmLogActaConformidadNew 'PASI20140925 ERS0772014
            oActa.Inicio gnAlmaActaConformidadME, "ACTA DE CONFORMIDAD DIGITAL DOLARES"
        Case 2:
            'Set oLibre = New frmLogActaConformidadLibre 'Comentado PASI20140925 ERS0772014
            Set oActa = New frmLogActaConformidadNew 'PASI20140925 ERS0772014
            'oLibre.Inicio gnAlmaActaConformidadLibreMN, "ACTA DE CONFORMIDAD DIGITAL LIBRE SOLES" 'Comentado PASI20140925 ERS0772014
            oActa.Inicio gnAlmaActaConformidadLibreMN, "ACTA DE CONFORMIDAD DIGITAL LIBRE SOLES"
        Case 3:
            'Set oLibre = New frmLogActaConformidadLibre 'Comentado PASI20140925 ERS0772014
            Set oActa = New frmLogActaConformidadNew 'PASI20140925 ERS0772014
            'oLibre.Inicio gnAlmaActaConformidadLibreME, "ACTA DE CONFORMIDAD DIGITAL LIBRE DOLARES" 'Comentado PASI20140925 ERS0772014
            oActa.Inicio gnAlmaActaConformidadLibreME, "ACTA DE CONFORMIDAD DIGITAL LIBRE DOLARES"
        Case 4:
            Set oExtorno = New frmLogActaConformidadExtorno
            oExtorno.Show 1
    End Select
    Set oExtorno = Nothing
    'Set oLibre = Nothing
    Set oActa = Nothing
End Sub
Private Sub M1701050800_Click(index As Integer)
    Select Case index
        Case 0:
            frmLogComprobanteRegistro.Show 1
        Case 2: 'PASIERS1242014 cambio 1 x 2
            frmLogImpComprobantes.Inicio
        Case 3: 'PASIERS1242014 cambio 2 x 3
            frmLogComprobanteHistorial.Show 1
    End Select
End Sub

'END EJVG *******
'WIOR FIN **************************************
Private Sub M1701070000_Click(index As Integer)
    If index = 0 Then
        frmLogOperacionesIngBS.Inicio "59"
    ElseIf index = 1 Then
        frmLogAlmInven.Show 1
    ElseIf index = 2 Then
        frmLogKardex.Show 1
    ElseIf index = 3 Then
        frmLogCalculaSaldos.Show 1
    ElseIf index = 4 Then
        frmLogMantCtaCont.Show 1
    ElseIf index = 5 Then
        frmLogMantSaldos.Show 1
    ElseIf index = 6 Then
        frmLogCalculaAsientos.Show 1
    
    ElseIf index = 10 Then
        frmLogEstadisticas.Show 1
    ElseIf index = 11 Then
        frmLogEstadConsumo.Show 1
    ElseIf index = 12 Then
        frmLogEstadAtencion.Show 1
        'frmControles.Show 1
       
    End If
End Sub

Private Sub M1701080000_Click(index As Integer)
    If index = 0 Then
        frmLogAFDeprecia.Ini "RECURSOS HUMANOS:DEPRECIACION DE ACTIVO FIJO"
    ElseIf index = 2 Then
        'frmTransferencia.Show 1
        frmLogBienTransferencia.Show 1 'EJVG20130626
    ElseIf index = 3 Then
        frmAFAsiganPersona.Show 1
    ElseIf index = 4 Then
        frmAFBajaActivo.Show 1
    ElseIf index = 6 Then
        frmReporteMovBienes.Show 1
    ElseIf index = 7 Then
        frmReporteKardexActivo.Show 1
    ElseIf index = 8 Then
        'frmModifyActivo.Show 1
        frmLogBienMnt.Show 1 'EJVG20130626
    ElseIf index = 9 Then
        frmLogBienAjusteVidaUtil.Show 1 'EJVG20130621
    ElseIf index = 10 Then
        frmLogBienBajaDestino.Show 1 'EJVG20130621
    ElseIf index = 11 Then
        frmLogBienDeterioro.Show 1 'EJVG20130621
    End If
End Sub

Private Sub M1701090000_Click(index As Integer)
    If index = 0 Then 'Transferencia
    
    ElseIf index = 1 Then 'Asignacion
        frmBNDAsiganPersona.Show 1
    ElseIf index = 2 Then 'Baja
        frmBNDBaja.Show 1
    End If
End Sub

Private Sub M1701100000_Click(index As Integer)
    '1 - REGISTRO ;     2 - DISTRIBUCION;     3 - GARANTIA
    frmLogSerCon.Inicio (index + 1)
End Sub

Private Sub M1801000000_Click(index As Integer)
    If index = 0 Then
        frmPreMantenimiento.Show 1
    ElseIf index = 1 Then
        frmPreRubros.Show 1
    ElseIf index = 2 Then
        frmPlaPresu.Show 1
    ElseIf index = 3 Then
        frmPlaEjecu.Show 1
    End If
End Sub

Private Sub M1801010000_Click(index As Integer)
    If index = 0 Then
        frmLogOCAtencion.Inicio True, "501221", True
    Else
        frmLogOCAtencion.Inicio True, "502221", True
    End If
    Set frmLogOCAtencion = Nothing
End Sub

Private Sub M1801020000_Click(index As Integer)
    If index = 0 Then
        frmLogOCAtencion.Inicio False, "501222", True
    Else
        frmLogOCAtencion.Inicio False, "502222", True
    End If
    Set frmLogOCAtencion = Nothing
End Sub

Private Sub M1901010000_Click(index As Integer)
    If index = 0 Then
        frmRHSeleccion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuResultado, "SIG:RECURSOS HUMANOS:PROCESO SELECCION:CONSULTA"
    ElseIf index = 2 Then
        frmRHEmpleado.Ini gTipoOpeConsulta, RHContratoMantTpoFoto, "SIG:RECURSOS HUMANOS:FICHA PERSONAL:CONSULTA"
    ElseIf index = 4 Then
        frmRHAsistenciaManual.Ini gTipoOpeConsulta, "SIG:RECURSOS HUMANOS:ASISTENCIA:MANUAL:CONSULTA"
    ElseIf index = 6 Then
        frmRHPlanillas.Ini gTipoProcesoRRHHConsulta, "RECURSOS HUMANOS:PROCESOS:ABONO DE PLANILLAS", Me
    ElseIf index = 8 Then
        frmRRHHRep.Ini "RECURSOS HUMANOS:REPORTES:REPORTES", Me
    ElseIf index = 9 Then
        frmRRHHRepGen.Ini "RECURSOS HUMANOS:REPORTES:REPORTES GENERALES", Me
    ElseIf index = 10 Then
        frmRHRepPresupuesto.Show
        
    End If
End Sub

Private Sub M1901020000_Click(index As Integer)
    If index = 0 Then
        frmLogMantSaldos.Inicio True
    ElseIf index = 2 Then
        frmLogOperacionesIngBS.Inicio "5915"
    End If
End Sub

Private Sub M1901030000_Click(index As Integer)
     If index = 0 Then
        frmPlaEjecu.Show 1
    End If
End Sub

Private Sub M2001010000_Click(index As Integer)
    'frmInvActivarBien.Show
    frmInvActivarBien_NEW.Inicio gMonedaNacional 'EJVG20130617
End Sub

Private Sub M2001020000_Click(index As Integer)
    'frmInvActivarBienDolares.Show
    frmInvActivarBien_NEW.Inicio gMonedaExtranjera 'EJVG20130617
End Sub

Private Sub M2002000000_Click(index As Integer)
    frmInvTransferenciaBS.Show
End Sub

Private Sub M2003010000_Click(index As Integer)
    frmInvReporteAF.Show 1
End Sub


Private Sub M2003020000_Click(index As Integer)
    frmInvReporteAsientoT.Show
End Sub

Private Sub M2003030000_Click(index As Integer)
    frmInvReporteTransferencia.Show
End Sub

'Private Sub M2004000000_Click(Index As Integer)
'    frmLogBienMnt.Show 1
'End Sub
'
'Private Sub M2005000000_Click(Index As Integer)
'    frmLogBienTransferencia.Show 1
'End Sub

'EJVG20120925 ***
Private Sub M2101000000_Click(index As Integer)
    Select Case index
        Case 1
            frmMktProdServ.Show 1
        Case 2
            frmMktActividad.Show 1
        Case 3
            frmMktCompras.Show 1
        Case 4
            frmMktUsoProdServ.Show 1
        'VAPI SEGUN ERS-0822014
        Case 5
            frmMkConfCombosCampana.Show 1
        Case 6
            frmMKEntregasDirecta.Show 1
        'END VAPI
        
    End Select
End Sub
Private Sub M2101010000_Click(index As Integer)
    Select Case index
        Case 0
            frmMktParametrosGastos.Inicio (1)
        Case 1
            frmMktParametrosGastos.Inicio (2)
        Case 2
            frmMktParametrosGastos.Inicio (3)
    End Select
End Sub
'END EJVG *******
Private Sub MDIForm_Load()
'    Dim Cont As Control
'    On Error Resume Next
'    For Each Cont In Controls ' Itera por cada elemento.
'        If Left(Cont.Name, 1) = "M" Then
'            If Left(Cont.Name, 3) <> "M09" And Left(Cont.Name, 3) <> "M18" And Left(Cont.Name, 3) <> "M17" And Left(Cont.Name, 3) <> "M01" And Left(Cont.Name, 3) <> "M16" Then
'            'If Left(Cont.Name, 3) <> "M09" And Left(Cont.Name, 3) <> "M18" And Left(Cont.Name, 3) <> "M17" And Left(Cont.Name, 3) <> "M01" Then
'              'Cont.Visible = False
'            End If
'        End If
'    Next
 '********* Temporalmente ***********
    Dim R As New ADODB.Recordset
    Dim oConec As DConecta
    Dim sSQL As String
    Dim Y As Integer
    
    '->***** LUCV20190323, Según RO-1000373
    'Quita el borde de los dos controles
    Image1.BorderStyle = 0
    pbxFondo.BorderStyle = 0
    '->***** Fin LUCV20190323
 
 
 '   gsRutaIcono = "\cm.ico"
    'Timer1.Enabled = False
    'NroRegOpe = 1628
 '   gdFecSis = Date
 '   Vusuario = "ARCV"
 '   gsCodAge = "01"
    gsCodCMAC = "109"

    'Habilita Permiso para Operaciones y Reportes
    Set oConec = New DConecta
    oConec.AbreConexion
    sSQL = "Select cOpeCod,cOpeDesc,cOpeVisible,nOpeNiv,cOpeGruCod,cUltimaActualizacion from OpeTpo WHERE cOpeVisible ='1' Order by cOpeCod"
    Set R = oConec.CargaRecordSet(sSQL)
    Y = 0
    Do While Not R.EOF
        Y = Y + 1
        MatOperac(Y - 1, 0) = R!cOpeCod
        MatOperac(Y - 1, 1) = R!cOpeDesc
        MatOperac(Y - 1, 2) = IIf(IsNull(R!cOpeGruCod), "", R!cOpeGruCod)
        MatOperac(Y - 1, 3) = R!cOpeVisible
        MatOperac(Y - 1, 4) = R!nOpeNiv
        R.MoveNext
    Loop
    NroRegOpe = Y
    oConec.CierraConexion
    CargaMensajes 'WIOR 20130826
    
 '************************************
 
 '->***** LUCV20190323, Según RO-1000373
    If VerificaGrupoMantenimientoUsuarios Then
        tlbMain.Buttons.item(2).Visible = True
        tlbMain.Buttons.item(2).Enabled = True
    Else
        tlbMain.Buttons.item(2).Visible = False
        tlbMain.Buttons.item(2).Enabled = False
    End If
 '<-***** Fin LUCV20190323
End Sub


Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > 3650 And X < 3740 And Y > 0 And Y < 50 Then
        If Shift = 7 And Button = 2 Then
            'frmImagenTransferencia.Show 1
        End If
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    'COOMENTADO POR ARLO 20170511
    'ALPA 20090122 ***************************************************************************
'     Set objPista = New COMManejador.Pista
'     glsMovNro = GetMovNro(gsCodUser, gsCodAge)
'     gsOpeCod = LogPistaIngresarSalirSistema 'COMENTADO POR ARLO20170511
'     objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, gSalirSistema
    '*****************************************************************************************
    '***********************
    'ARLO 20170511
    If MsgBox("Esta Seguro que Desea Salir del Sistema ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
    Call SalirSICMACMAdmnstrativo
    End If
    '*********
End Sub
'WIOR 20130826 *************************************************************
Private Sub CargaMensajes()
Dim oSeg As UAcceso
Dim rsSeg As ADODB.Recordset

Set oSeg = New UAcceso
Set rsSeg = oSeg.ObtenerMensajeSeguridad()

If Not (rsSeg.EOF And rsSeg.BOF) Then
    frmSegMensajeMostrar.Inicio (Trim(rsSeg!cMensaje))
End If

Set oSeg = Nothing
Set rsSeg = Nothing
End Sub
'WIOR FIN *******************************************************************
'PASIERS0772014
Private Sub M1701050801_Click(index As Integer)
    Select Case index
        Case 0
                frmLogComprobanteRegistroNew.Inicio gnAlmaComprobanteRegistroMN
        Case 1
                frmLogComprobanteRegistroNew.Inicio gnAlmaComprobanteRegistroME
        Case 2
                frmLogComprobanteLibreRegistro.Inicio gnAlmaComprobanteLibreRegistroMN
        Case 3
                frmLogComprobanteLibreRegistro.Inicio gnAlmaComprobanteLibreRegistroME
    End Select
End Sub
Private Sub M1701050804_Click(index As Integer)
    Select Case index
        Case 0
                frmLogComprobanteExtorno.Inicio gnAlmaComprobanteExtornoMN
        Case 1
                frmLogComprobanteExtorno.Inicio gnAlmaComprobanteExtornoME
        Case 2
                frmLogComprobanteExtorno.Inicio gnAlmaComprobanteLibreExtornoMN
        Case 3
                frmLogComprobanteExtorno.Inicio gnAlmaComprobanteLibreExtornoME
    End Select
End Sub
'END PASI

'ARLO 20170511
Sub SalirSICMACMAdmnstrativo()
    Set objPista = New COMManejador.Pista 'LogPistaIngresoSistema
    Call objPista.InsertarPista(LogPistaIngresoSistema, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gSalirSistema, "Salida del Sicmac Administrativo" & " Versión: " & Format(App.Major, "#0") & "." & Format(App.Minor, "#0") & "." & Format(App.Revision, "#0") & "-" & gsFechaVersion)
    Set objPista = Nothing
End Sub

'->***** LUCV20190323, Según RO-1000373
Public Sub SCalculadora()
    Dim valor, I
    valor = Shell("calc.exe", 1)  ' Ejecuta la Calculadora.
End Sub

Private Function VerificaGrupoMantenimientoUsuarios() As Boolean
    Dim oCons As NConstSistemas
    Set oCons = New NConstSistemas
    Dim sGrupoAutorizado As String
    Dim nGrupoTmp1 As String
    Dim nGrupoTmp2 As String
    Dim I As Integer
    Dim J As Integer
            
    sGrupoAutorizado = oCons.LeeConstSistema(519)
    VerificaGrupoMantenimientoUsuarios = False
    For I = 1 To Len(sGrupoAutorizado)
        If Not Mid(sGrupoAutorizado, I, 1) = "," Then
            nGrupoTmp1 = nGrupoTmp1 & Mid(sGrupoAutorizado, I, 1)
        Else
            For J = 1 To Len(gsGrupoUsu)
                If Not Mid(gsGrupoUsu, J, 1) = "," Then
                    nGrupoTmp2 = nGrupoTmp2 & Mid(gsGrupoUsu, J, 1)
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

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.index
        Case 1
            SCalculadora
        Case 2
            frmMantPermisos.Show 1
    End Select
    
    Exit Sub
ErrorBoton:
End Sub
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
