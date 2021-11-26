VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDISicmact 
   BackColor       =   &H8000000C&
   Caption         =   "SICMACT Sistema Integrado de la Caja Municipal de Ahorro y Credito de Trujillo"
   ClientHeight    =   6510
   ClientLeft      =   135
   ClientTop       =   1845
   ClientWidth     =   11625
   Icon            =   "MdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar staMain 
      Align           =   2  'Align Bottom
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   6285
      Width           =   11625
      _ExtentX        =   20505
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
      Left            =   1050
      Top             =   1350
   End
   Begin VB.Menu M0100000000 
      Caption         =   "Arc&hivo"
      Index           =   0
      Begin VB.Menu M0101000000 
         Caption         =   "Configurar &Impresora"
         Index           =   0
      End
      Begin VB.Menu M0101000000 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu M0101000000 
         Caption         =   "&Salir"
         Index           =   2
      End
   End
   Begin VB.Menu M0200000000 
      Caption         =   "&Creditos"
      Index           =   0
      Begin VB.Menu M0201000000 
         Caption         =   "&Definiciones"
         Index           =   0
         Begin VB.Menu M0201010000 
            Caption         =   "&Parametros de Control"
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
            Caption         =   "&Lineas de Credito"
            Index           =   1
            Begin VB.Menu M0201010200 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M0201010200 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M0201010200 
               Caption         =   "&Consulta"
               Index           =   3
            End
         End
         Begin VB.Menu M0201010000 
            Caption         =   "&Niveles de Aprobacion"
            Index           =   2
            Begin VB.Menu M0201010300 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0201010300 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0201010000 
            Caption         =   "&Gastos"
            Index           =   3
            Begin VB.Menu M0201010400 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0201010400 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Solicitud Credito"
         Index           =   1
         Begin VB.Menu M0201020000 
            Caption         =   "&Registro"
            Index           =   0
         End
         Begin VB.Menu M0201020000 
            Caption         =   "&Consulta"
            Index           =   1
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Relaciones de Credito"
         Index           =   2
         Begin VB.Menu M0201030000 
            Caption         =   "&Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu M0201030000 
            Caption         =   "Reasignacion de &Cartera en Lote"
            Index           =   1
         End
         Begin VB.Menu M0201030000 
            Caption         =   "Con&sulta"
            Index           =   2
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Garantias"
         Index           =   3
         Begin VB.Menu M0201040000 
            Caption         =   "&Registro"
            Index           =   0
         End
         Begin VB.Menu M0201040000 
            Caption         =   "&Mantenimiento"
            Index           =   1
         End
         Begin VB.Menu M0201040000 
            Caption         =   "&Consulta"
            Index           =   2
         End
         Begin VB.Menu M0201040000 
            Caption         =   "Gravamen"
            Index           =   3
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Sugerencia"
         Index           =   4
         Begin VB.Menu M0201050000 
            Caption         =   "&Registro"
            Index           =   0
         End
         Begin VB.Menu M0201050000 
            Caption         =   "&Mantenimiento"
            Index           =   1
         End
         Begin VB.Menu M0201050000 
            Caption         =   "&Consulta"
            Index           =   2
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Resolver Creditos"
         Index           =   5
         Begin VB.Menu M0201060000 
            Caption         =   "&Aprobacion"
            Index           =   0
         End
         Begin VB.Menu M0201060000 
            Caption         =   "&Rechazo"
            Index           =   1
         End
         Begin VB.Menu M0201060000 
            Caption         =   "A&nulacion Creditos Aprobados"
            Index           =   2
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Reprogramacion Credito"
         Index           =   6
         Begin VB.Menu M0201070000 
            Caption         =   "Repr&ogramacion"
            Index           =   0
         End
         Begin VB.Menu M0201070000 
            Caption         =   "Reprogramacion en &Lote"
            Index           =   1
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Refinanciacion"
         Index           =   7
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Actualizacion de &Metodos de Liquidacion"
         Index           =   8
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Perdonar Mora"
         Index           =   9
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Gastos"
         Index           =   10
         Begin VB.Menu M0201080000 
            Caption         =   "&Administracion de Gastos en &Lote"
            Index           =   0
         End
         Begin VB.Menu M0201080000 
            Caption         =   "Mantenimiento de Penalidad de Cancelacion"
            Index           =   1
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Reasignar &Institucion"
         Index           =   11
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Transferencia a Recuperaciones"
         Index           =   12
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Analista"
         Index           =   13
         Begin VB.Menu M0201090000 
            Caption         =   "&Nota de Analista"
            Index           =   0
         End
         Begin VB.Menu M0201090000 
            Caption         =   "&Metas de Analista"
            Index           =   1
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Si&mulaciones"
         Index           =   14
         Begin VB.Menu M0201100000 
            Caption         =   "Calendario de &Pagos"
            Index           =   0
         End
         Begin VB.Menu M0201100000 
            Caption         =   "Calendario de &Desembolsos Parciales"
            Index           =   1
         End
         Begin VB.Menu M0201100000 
            Caption         =   "Calendario de Cuota &Libre"
            Index           =   2
         End
         Begin VB.Menu M0201100000 
            Caption         =   "Simulador de Pagos"
            Index           =   3
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Cons&ultas"
         Index           =   15
         Begin VB.Menu M0201200000 
            Caption         =   "&Historial de Credito"
            Index           =   0
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Reportes"
         Index           =   16
         Begin VB.Menu M0201300000 
            Caption         =   "&Duplicados"
            Index           =   0
         End
      End
   End
   Begin VB.Menu M0300000000 
      Caption         =   "&Pignoraticio"
      Index           =   0
      Begin VB.Menu M0301000000 
         Caption         =   "&Contratro"
         Index           =   0
         Begin VB.Menu M0301010000 
            Caption         =   "&Registro"
            Index           =   0
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Mantenimiento Descripcion"
            Index           =   1
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Anulación"
            Index           =   2
         End
         Begin VB.Menu M0301010000 
            Caption         =   "&Bloqueo"
            Index           =   3
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "Rescate Joyas"
         Index           =   2
      End
      Begin VB.Menu M0301000000 
         Caption         =   "Remate"
         Index           =   3
         Begin VB.Menu M0301020000 
            Caption         =   "Preparacion Remate"
            Index           =   0
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Remate"
            Index           =   1
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "Adjudicacion"
         Index           =   4
      End
      Begin VB.Menu M0301000000 
         Caption         =   "Subasta"
         Index           =   5
         Begin VB.Menu M0301030000 
            Caption         =   "Preparacion Subasta"
            Index           =   0
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Subasta"
            Index           =   1
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "&Reportes"
         Index           =   6
         Begin VB.Menu M0301040000 
            Caption         =   "Movimientos Diario"
            Index           =   0
         End
         Begin VB.Menu M0301040000 
            Caption         =   "Listados Generales"
            Index           =   1
         End
         Begin VB.Menu M0301040000 
            Caption         =   "Estadisticas"
            Index           =   2
         End
         Begin VB.Menu M0301040000 
            Caption         =   "Contrato de Persona"
            Index           =   3
         End
      End
   End
   Begin VB.Menu M0400000000 
      Caption         =   "&Captaciones"
      Index           =   0
      Begin VB.Menu M0401000000 
         Caption         =   "&Parámetros"
         Index           =   0
         Begin VB.Menu M0401010000 
            Caption         =   "&Consulta"
            Index           =   0
         End
         Begin VB.Menu M0401010000 
            Caption         =   "&Mantenimiento"
            Index           =   1
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Tasas de Interés"
         Index           =   1
         Begin VB.Menu M0401020000 
            Caption         =   "&Ahorros"
            Index           =   0
            Begin VB.Menu M0401020100 
               Caption         =   "&Consulta"
               Index           =   0
            End
            Begin VB.Menu M0401020100 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
         End
         Begin VB.Menu M0401020000 
            Caption         =   "&Plazo Fijo"
            Index           =   1
            Begin VB.Menu M0401020200 
               Caption         =   "&Consulta"
               Index           =   0
            End
            Begin VB.Menu M0401020200 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
         End
         Begin VB.Menu M0401020000 
            Caption         =   "&CTS"
            Index           =   2
            Begin VB.Menu M0401020300 
               Caption         =   "&Consulta"
               Index           =   0
            End
            Begin VB.Menu M0401020300 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Mantenimiento"
         Index           =   2
         Begin VB.Menu M0401030000 
            Caption         =   "&Ahorros"
            Index           =   0
         End
         Begin VB.Menu M0401030000 
            Caption         =   "&Plazo Fijo"
            Index           =   1
         End
         Begin VB.Menu M0401030000 
            Caption         =   "&CTS"
            Index           =   2
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Bloqueo/Desbloqueo"
         Index           =   3
         Begin VB.Menu M0401040000 
            Caption         =   "&Ahorros"
            Index           =   0
         End
         Begin VB.Menu M0401040000 
            Caption         =   "&Plazo Fijo"
            Index           =   1
         End
         Begin VB.Menu M0401040000 
            Caption         =   "&CTS"
            Index           =   2
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Simulación Plazo Fijo"
         Index           =   4
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Tarjeta Magnética"
         Index           =   5
         Begin VB.Menu M0401050000 
            Caption         =   "&Registro"
            Index           =   0
         End
         Begin VB.Menu M0401050000 
            Caption         =   "Re&lación"
            Index           =   1
         End
         Begin VB.Menu M0401050000 
            Caption         =   "&Bloqueo/Desbloqueo"
            Index           =   2
         End
         Begin VB.Menu M0401050000 
            Caption         =   "&Cambio de Clave"
            Index           =   3
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Bene&ficiarios"
         Index           =   6
         Begin VB.Menu M0401060000 
            Caption         =   "&Consulta"
            Index           =   0
         End
         Begin VB.Menu M0401060000 
            Caption         =   "&Mantenimiento"
            Index           =   1
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Orden Pago"
         Index           =   7
         Begin VB.Menu M0401070000 
            Caption         =   "&Generación y Emisión"
            Index           =   0
         End
         Begin VB.Menu M0401070000 
            Caption         =   "&Certificación"
            Index           =   1
         End
         Begin VB.Menu M0401070000 
            Caption         =   "&Anulación"
            Index           =   2
         End
         Begin VB.Menu M0401070000 
            Caption         =   "Con&sulta"
            Index           =   3
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Ser&vicios"
         Index           =   8
         Begin VB.Menu M0401080000 
            Caption         =   "&Sedalib"
            Index           =   0
            Begin VB.Menu M0401080100 
               Caption         =   "&Parámetros"
               Index           =   0
            End
            Begin VB.Menu M0401080100 
               Caption         =   "&Reportes"
               Index           =   1
            End
            Begin VB.Menu M0401080100 
               Caption         =   "&Generación DBF"
               Index           =   2
            End
         End
         Begin VB.Menu M0401080000 
            Caption         =   "&Hidrandina"
            Index           =   1
            Begin VB.Menu M0401080200 
               Caption         =   "&Parámetros"
               Index           =   0
            End
            Begin VB.Menu M0401080200 
               Caption         =   "&Reportes"
               Index           =   1
            End
         End
         Begin VB.Menu M0401080000 
            Caption         =   "&Universidad Nacional de Trujillo"
            Index           =   2
            Begin VB.Menu M0401080300 
               Caption         =   "&Parámetros"
               Index           =   0
            End
            Begin VB.Menu M0401080300 
               Caption         =   "&Reportes"
               Index           =   1
            End
            Begin VB.Menu M0401080300 
               Caption         =   "&Generación DBF"
               Index           =   2
            End
         End
         Begin VB.Menu M0401080000 
            Caption         =   "&Instituciones Convenio"
            Index           =   3
            Begin VB.Menu M0401080400 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0401080400 
               Caption         =   "&Cuentas"
               Index           =   1
            End
            Begin VB.Menu M0401080400 
               Caption         =   "&Plan de Pagos"
               Index           =   2
            End
         End
      End
   End
   Begin VB.Menu M0500000000 
      Caption         =   "&Recuperaciones"
      Index           =   0
      Begin VB.Menu M0501000000 
         Caption         =   "&Comisiones Abogados"
         Index           =   0
      End
      Begin VB.Menu M0501000000 
         Caption         =   "&Registro de Otros Creditos "
         Index           =   1
      End
      Begin VB.Menu M0501000000 
         Caption         =   "Proceso &Judicial"
         Index           =   2
         Begin VB.Menu M0501010000 
            Caption         =   "&Expedientes Judiciales"
            Index           =   0
         End
         Begin VB.Menu M0501010000 
            Caption         =   "Actuaciones &Procesales"
            Index           =   1
         End
      End
      Begin VB.Menu M0501000000 
         Caption         =   "&Gastos de Recuperacines"
         Index           =   3
      End
      Begin VB.Menu M0501000000 
         Caption         =   "Metodo de &Liquidacion"
         Index           =   4
      End
      Begin VB.Menu M0501000000 
         Caption         =   "Cas&tigar Credito"
         Index           =   5
      End
      Begin VB.Menu M0501000000 
         Caption         =   "Pago de Creditos"
         Index           =   6
      End
      Begin VB.Menu M0501000000 
         Caption         =   "Extorno de Pagos"
         Index           =   7
      End
      Begin VB.Menu M0501000000 
         Caption         =   "Consulta Para&metros"
         Index           =   8
      End
      Begin VB.Menu M0501000000 
         Caption         =   "Mantenimiento Parametros"
         Index           =   9
      End
   End
   Begin VB.Menu M0600000000 
      Caption         =   "&Operaciones"
      Index           =   0
      Begin VB.Menu M0601000000 
         Caption         =   "Desembolso &Abono a Cuenta"
         Index           =   0
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Operaciones"
         Index           =   1
         Shortcut        =   {F2}
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Operaciones CMACs &Recepción"
         Index           =   2
         Shortcut        =   {F3}
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Operaciones CMACs &Llamada"
         Index           =   3
         Shortcut        =   {F4}
      End
      Begin VB.Menu M0601000000 
         Caption         =   "&Extornos"
         Index           =   4
         Begin VB.Menu M0601010000 
            Caption         =   "Extornos de Credito"
            Index           =   0
         End
         Begin VB.Menu M0601010000 
            Caption         =   "Extornos de Pignoraticio"
            Index           =   1
         End
         Begin VB.Menu M0601010000 
            Caption         =   "Extornos de Recuperaciones"
            Index           =   2
         End
         Begin VB.Menu M0601010000 
            Caption         =   "Extornos de Captaciones"
            Index           =   3
         End
      End
   End
   Begin VB.Menu M0700000000 
      Caption         =   "Perso&nas"
      Index           =   0
      Begin VB.Menu M0701000000 
         Caption         =   "&Personas"
         Index           =   0
         Begin VB.Menu M0701010000 
            Caption         =   "&Registro"
            Index           =   0
         End
         Begin VB.Menu M0701010000 
            Caption         =   "&Manteniemiento"
            Index           =   1
         End
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
         Caption         =   "&Asignar Permisos"
         Index           =   0
      End
      Begin VB.Menu M1001000000 
         Caption         =   "Administracion de &Usuario"
         Index           =   1
      End
      Begin VB.Menu M1001000000 
         Caption         =   "Prueba3"
         Index           =   2
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
               Caption         =   "&Registro"
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
         Caption         =   "&Descansos"
         Index           =   15
         Begin VB.Menu M1601090000 
            Caption         =   "&Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu M1601090000 
            Caption         =   "&Consulta"
            Index           =   1
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
            Caption         =   "Cierre &Mensual"
            Index           =   2
         End
         Begin VB.Menu M1601190000 
            Caption         =   "Cierre &Diario"
            Index           =   3
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
      End
   End
   Begin VB.Menu M1700000000 
      Caption         =   "&Logística"
      Index           =   0
      Begin VB.Menu M1701000000 
         Caption         =   "Usuarios"
         Index           =   0
      End
      Begin VB.Menu M1701000000 
         Caption         =   "-"
         Index           =   1
      End
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
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Requerimiento Regular"
         Index           =   6
         Begin VB.Menu M1701010000 
            Caption         =   "&Solicitud"
            Index           =   0
         End
         Begin VB.Menu M1701010000 
            Caption         =   "&Trámite"
            Index           =   1
         End
         Begin VB.Menu M1701010000 
            Caption         =   "Precios Referenciales"
            Index           =   2
         End
         Begin VB.Menu M1701010000 
            Caption         =   "Aprobación o Rechazo"
            Index           =   3
         End
         Begin VB.Menu M1701010000 
            Caption         =   "Asignación de Ctas.Presupuestrales"
            Index           =   4
         End
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Requerimiento &Extemporaneo"
         Index           =   7
         Begin VB.Menu M1701020000 
            Caption         =   "Solicitud"
            Index           =   0
         End
         Begin VB.Menu M1701020000 
            Caption         =   "Trámite"
            Index           =   1
         End
         Begin VB.Menu M1701020000 
            Caption         =   "Precio Referencial"
            Index           =   2
         End
         Begin VB.Menu M1701020000 
            Caption         =   "Asignación de Ctas.Presupuestrales"
            Index           =   3
         End
         Begin VB.Menu M1701020000 
            Caption         =   "Aprobacion o Rechazo"
            Index           =   4
         End
      End
      Begin VB.Menu M1701000000 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Plan Anual de Adquisiciones y Contrataciones"
         Index           =   9
         Begin VB.Menu M1701030000 
            Caption         =   "Consolidación de los Requerimientos al plan Anual de Adquisiciones "
            Index           =   0
         End
         Begin VB.Menu M1701030000 
            Caption         =   "Consultas"
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
         Begin VB.Menu M1701040000 
            Caption         =   "Definición de Parámetros"
            Index           =   0
         End
         Begin VB.Menu M1701040000 
            Caption         =   "Inicio de proceso"
            Index           =   1
            Begin VB.Menu M1701040100 
               Caption         =   "Resolución o acuerdo de inicio"
               Index           =   0
               Begin VB.Menu M1701040101 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu M1701040101 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu M1701040100 
               Caption         =   "Conformación del Comite especial u organo encargado"
               Index           =   1
               Begin VB.Menu M1701040102 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu M1701040102 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu M1701040100 
               Caption         =   "Detalle del proceso"
               Index           =   2
               Begin VB.Menu M1701040103 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu M1701040103 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu M1701040100 
               Caption         =   "Bases o terminos de referencia"
               Index           =   3
               Begin VB.Menu M1701040104 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu M1701040104 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
         End
         Begin VB.Menu M1701040000 
            Caption         =   "Convocatoria"
            Index           =   2
            Begin VB.Menu M1701040200 
               Caption         =   "Publicación"
               Index           =   0
               Begin VB.Menu M1701040105 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu M1701040105 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu M1701040200 
               Caption         =   "Solicitudes de cotización"
               Index           =   1
               Begin VB.Menu M1701040106 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu M1701040106 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
         End
         Begin VB.Menu M1701040000 
            Caption         =   "Bases"
            Index           =   3
            Begin VB.Menu M1701040300 
               Caption         =   "Registro de entrega de Bases"
               Index           =   0
               Begin VB.Menu M1701040301 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu M1701040301 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu M1701040300 
               Caption         =   "Consulta a Bases"
               Index           =   1
               Begin VB.Menu M1701040302 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu M1701040302 
                  Caption         =   "Cosulta"
                  Index           =   1
               End
            End
            Begin VB.Menu M1701040300 
               Caption         =   "Absolución de Consultas"
               Index           =   2
               Begin VB.Menu M1701040303 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu M1701040303 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu M1701040300 
               Caption         =   "Registro de Observaciones"
               Index           =   3
               Begin VB.Menu M1701040304 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu M1701040304 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
         End
         Begin VB.Menu M1701040000 
            Caption         =   "Propuestas"
            Index           =   4
            Begin VB.Menu M1701040400 
               Caption         =   "Técnica"
               Index           =   0
            End
            Begin VB.Menu M1701040400 
               Caption         =   "Económica"
               Index           =   1
            End
            Begin VB.Menu M1701040400 
               Caption         =   "Garantía de seriedad de oferta"
               Index           =   2
            End
         End
         Begin VB.Menu M1701040000 
            Caption         =   "Evaluaciones"
            Index           =   5
         End
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Proceso de Selección Desierto"
         Index           =   12
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Cancelacion Proceso Seleccion"
         Index           =   13
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Adjudicación del Proceso de Selección"
         Index           =   14
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Consentimiento de Adjudicación"
         Index           =   15
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Consultas"
         Index           =   16
      End
      Begin VB.Menu M1701000000 
         Caption         =   "-"
         Index           =   17
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
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Provisiones pago de Proveedores"
         Index           =   19
         Begin VB.Menu M1701060000 
            Caption         =   "Ordenes de Compra Soles"
            Index           =   0
         End
         Begin VB.Menu M1701060000 
            Caption         =   "Ordenes de Servicios Soles"
            Index           =   1
         End
         Begin VB.Menu M1701060000 
            Caption         =   "Ordenes de Compra Dolares"
            Index           =   2
         End
         Begin VB.Menu M1701060000 
            Caption         =   "Ordenes de Servicios Dolares"
            Index           =   3
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
            Caption         =   "Recalculo de Saldos"
            Index           =   3
         End
         Begin VB.Menu M1701070000 
            Caption         =   "Cierre de Dia"
            Index           =   4
         End
         Begin VB.Menu M1701070000 
            Caption         =   "Mantenimiento de Saldos"
            Index           =   5
         End
         Begin VB.Menu M1701070000 
            Caption         =   "Regeneraciòn de Asientos Contables"
            Index           =   6
         End
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Depreciación de Bienes"
         Index           =   22
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Subasta de Bienes"
         Index           =   23
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Transferencia de Activo Fijo"
         Index           =   24
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Remate"
         Index           =   25
      End
      Begin VB.Menu M1701000000 
         Caption         =   "-"
         Index           =   26
      End
      Begin VB.Menu M1701000000 
         Caption         =   "Servicios (locación, públicos, privados, móviles, otros)"
         Index           =   27
         Begin VB.Menu M1701080000 
            Caption         =   "Registro"
            Index           =   0
         End
         Begin VB.Menu M1701080000 
            Caption         =   "Distribución"
            Index           =   1
         End
         Begin VB.Menu M1701080000 
            Caption         =   "Garantías"
            Index           =   2
         End
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
End
Attribute VB_Name = "MDISicmact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub M0101000000_Click(Index As Integer)
    If Index = 0 Then
        frmImpresora.Show 1
    ElseIf Index = 2 Then
        Unload Me
    End If
End Sub


Private Sub M0701010000_Click(Index As Integer)
    Select Case Index
        Case 0
            frmPersona.Registrar
        Case 1
            frmPersona.Consultar
        Case 2
    End Select
End Sub

Private Sub M0901000000_Click(Index As Integer)
    If Index = 0 Then
        frmEditorSicmact.Show 1
    ElseIf Index = 1 Then 'Spooler
        frmSpooler.Show 1
    ElseIf Index = 3 Then 'Perifericos
    
    ElseIf Index = 5 Then 'BackUp
        
        frmBackUp.Show 1
    End If
    
End Sub

Private Sub M1001000000_Click(Index As Integer)
    Select Case Index
        Case 0
            frmMantPermisos.Show 1
        Case 1
            frmAdmUsu.Show 1
    End Select
End Sub

Private Sub M1601000000_Click(Index As Integer)
    If Index = 9 Then
        frmRHAsistenciaManual.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:ASISTENCIA:MANUAL"
    ElseIf Index = 23 Then
        frmRHEmpleado.Ini gTipoOpeRegistro, RHContratoMantTpoComentario, "RECURSOS HUMANOS:COMENTARIO:REGISTRO"
    End If
End Sub

Private Sub M1601010000_Click(Index As Integer)
    If Index = 3 Then
        frmRHSeleccion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuResultado, "RECURSOS HUMANOS:PROCESO SELECCION:RESULTADOS Y CIERRE"
    ElseIf Index = 4 Then
        frmRHSeleccion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuResultado, "RECURSOS HUMANOS:PROCESO SELECCION:CONSULTA"
    End If
End Sub

Private Sub M1601010100_Click(Index As Integer)
    If Index = 0 Then
        frmRHSeleccion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuSel, "RECURSOS HUMANOS:PROCESO SELECCION:REGISTRO"
    ElseIf Index = 1 Then
        frmRHSeleccion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuSel, "RECURSOS HUMANOS:PROCESO SELECCION:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHSeleccion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuSel, "RECURSOS HUMANOS:PROCESO SELECCION:CONSULTA"
    ElseIf Index = 3 Then
        frmRHSeleccion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuSel, "RECURSOS HUMANOS:PROCESO SELECCION:REPORTE"
    End If
End Sub

Private Sub M1601010200_Click(Index As Integer)
    If Index = 0 Then
        frmRHSeleccion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuPost, "RECURSOS HUMANOS:POSTULANTES:REGISTRO"
    ElseIf Index = 1 Then
        frmRHSeleccion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuPost, "RECURSOS HUMANOS:POSTULANTES:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHSeleccion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuPost, "RECURSOS HUMANOS:POSTULANTES:CONSULTA"
    ElseIf Index = 3 Then
        frmRHSeleccion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuPost, "RECURSOS HUMANOS:POSTULANTES:REPORTE"
    End If
End Sub

Private Sub M1601010301_Click(Index As Integer)
    If Index = 0 Then
        frmRHSeleccion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaCur, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:CURRICULAR:REGISTRO"
    ElseIf Index = 1 Then
        frmRHSeleccion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaCur, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:CURRICULAR:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHSeleccion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaCur, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:CURRICULAR:CONSULTA"
    ElseIf Index = 3 Then
        frmRHSeleccion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaCur, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:CURRICULAR:REPORTE"
    End If
End Sub

Private Sub M1601010302_Click(Index As Integer)
    If Index = 0 Then
        frmRHSeleccion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ESCRITO:REGISTRO"
    ElseIf Index = 1 Then
        frmRHSeleccion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ESCRITO:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHSeleccion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ESCRITO:CONSULTA"
    ElseIf Index = 3 Then
        frmRHSeleccion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ESCRITO:REPORTE"
    End If
End Sub

Private Sub M1601010303_Click(Index As Integer)
    If Index = 0 Then
        frmRHSeleccion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ESCRITO:REGISTRO"
    ElseIf Index = 1 Then
        frmRHSeleccion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ESCRITO:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHSeleccion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ESCRITO:CONSULTA"
    ElseIf Index = 3 Then
        frmRHSeleccion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ESCRITO:REPORTE"
    End If
End Sub

Private Sub M1601010304_Click(Index As Integer)
    If Index = 0 Then
        frmRHSeleccion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ENTREVISTA:REGISTRO"
    ElseIf Index = 1 Then
        frmRHSeleccion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ENTREVISTA:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHSeleccion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ENTREVISTA:CONSULTA"
    ElseIf Index = 3 Then
        frmRHSeleccion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ENTREVISTA:REPORTE"
    End If
End Sub

Private Sub M1601020000_Click(Index As Integer)
    If Index = 0 Then
        frmRHContratoSeleccion.Ini "RECURSOS HUMANOS:CONTRATO:PROCESO SELECCION", ContratoFormaAutomatica
    ElseIf Index = 1 Then
        frmRHContratoSeleccion.Ini "RECURSOS HUMANOS:CONTRATO:PROCESO SELECCION", ContratoFormaManual
    ElseIf Index = 2 Then
        frmRHEmpleado.Ini gTipoOpeMantenimiento, RHContratoMantTpoFoto, "RECURSOS HUMANOS:CONTRATOS:MANTENIMIENTO"
    ElseIf Index = 3 Then
        frmRHEmpleado.Ini gTipoOpeConsulta, RHContratoMantTpoFoto, "RECURSOS HUMANOS:FICHA PERSONAL:CONSULTA"
    ElseIf Index = 4 Then
        Me.Enabled = False
        frmRHEmpleadoResCont.Ini "RECURSOS HUMANOS:CONTRATO:RESCINDIR CONTRATO", Me
        Me.Enabled = True
    End If
End Sub

Private Sub M1601030000_Click(Index As Integer)
    frmRHEmpleado.Ini gTipoOpeRegistro, RHContratoMantTpoAdenda, "RECURSOS HUMANOS:ADENDA"
End Sub

Private Sub M1601040000_Click(Index As Integer)
    If Index = 1 Then
        frmRHCurriculum.Ini gTipoOpeRegistro, "RECURSOS HUMANOS:CURRICULUM VITAE:REGISTRO"
    ElseIf Index = 2 Then
        frmRHCurriculum.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:CURRICULUM VITAE:MANTENIMIENTO"
    ElseIf Index = 3 Then
        frmRHCurriculum.Ini gTipoOpeConsulta, "RECURSOS HUMANOS:CURRICULUM VITAE:CONSULTA"
    ElseIf Index = 4 Then
        frmRHCurriculum.Ini gTipoOpeReporte, "RECURSOS HUMANOS:CURRICULUM VITAE:REPORTE"
    End If
End Sub

Private Sub M1601040100_Click(Index As Integer)
    If Index = 0 Then
        frmRHCurriculumTabla.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:CURRICULUM VITAE:TABLA:MANTENIMIENTO"
    ElseIf Index = 1 Then
        frmRHCurriculumTabla.Ini gTipoOpeConsulta, "RECURSOS HUMANOS:CURRICULUM VITAE:TABLA:CONSULTA"
    End If
End Sub

Private Sub M1601050000_Click(Index As Integer)
    If Index = 0 Then
        frmRHAsistenciaAsig.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:HORARIO LABORAL:MANTENIMIENTO"
    ElseIf Index = 1 Then
        frmRHAsistenciaAsig.Ini gTipoOpeConsulta, "RECURSOS HUMANOS:HORARIO LABORAL:CONSULTA"
    End If
End Sub

Private Sub M1601060000_Click(Index As Integer)
    If Index = 5 Then
        frmRHEvaluacion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuResultado, "RECURSOS HUMANOS:EVA INT:RESULTADOS Y CIERRE"
    ElseIf Index = 6 Then
        frmRHEvaluacion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuResultado, "RECURSOS HUMANOS:EVA INT:CONSULTA"
    End If
End Sub

Private Sub M1601060100_Click(Index As Integer)
    If Index = 0 Then
        frmRHEvaluacion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuSel, "RECURSOS HUMANOS:EVA INT:EVALUACION:Proceso EVALUACION Interna:REGISTRO"
    ElseIf Index = 1 Then
        frmRHEvaluacion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuSel, "RECURSOS HUMANOS:EVA INT:EVALUACION:Proceso EVALUACION Interna:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHEvaluacion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuSel, "RECURSOS HUMANOS:EVA INT:EVALUACION:Proceso EVALUACION Interna:CONSULTA"
    ElseIf Index = 3 Then
        frmRHEvaluacion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuSel, "RECURSOS HUMANOS:EVA INT:EVALUACION:Proceso EVALUACION Interna:REPORTE"
    End If
End Sub

Private Sub M1601060200_Click(Index As Integer)
    If Index = 0 Then
        frmRHEvaluacion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaCur, "7RECURSOS HUMANOS:EVA INT:EVALUACION CURRICULAR:REGISTRO"
    ElseIf Index = 1 Then
        frmRHEvaluacion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaCur, "RECURSOS HUMANOS:EVA INT:EVALUACION CURRICULAR:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHEvaluacion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaCur, "RECURSOS HUMANOS:EVA INT:EVALUACION CURRICULAR:CONSULTA"
    ElseIf Index = 3 Then
        frmRHEvaluacion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaCur, "RECURSOS HUMANOS:EVA INT:EVALUACION CURRICULAR:REPORTE"
    End If
End Sub

Private Sub M1601060300_Click(Index As Integer)
    If Index = 0 Then
        frmRHEvaluacion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:EVA INT:EVALUACION Escrita:REGISTRO"
    ElseIf Index = 1 Then
        frmRHEvaluacion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:EVA INT:EVALUACION Escrita:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHEvaluacion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:EVA INT:EVALUACION Escrita:CONSULTA"
    ElseIf Index = 3 Then
        frmRHEvaluacion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:EVA INT:EVALUACION Escrita:REPORTE"
    End If
End Sub

Private Sub M1601060400_Click(Index As Integer)
    If Index = 0 Then
        frmRHEvaluacion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaPsi, "RECURSOS HUMANOS:EVA INT:EVALUACION Psicologica:REGISTRO"
    ElseIf Index = 1 Then
        frmRHEvaluacion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaPsi, "RECURSOS HUMANOS:EVA INT:EVALUACION Psicologica:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHEvaluacion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaPsi, "RECURSOS HUMANOS:EVA INT:EVALUACION Psicologica:CONSULTA"
    ElseIf Index = 3 Then
        frmRHEvaluacion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaPsi, "RECURSOS HUMANOS:EVA INT:EVALUACION Psicologica:REPORTE"
    End If
End Sub

Private Sub M1601060500_Click(Index As Integer)
    If Index = 0 Then
        frmRHEvaluacion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:EVA INT:EVALUACION Psicologica:REGISTRO"
    ElseIf Index = 1 Then
        frmRHEvaluacion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:EVA INT:EVALUACION Psicologica:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHEvaluacion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:EVA INT:EVALUACION Psicologica:CONSULTA"
    ElseIf Index = 3 Then
        frmRHEvaluacion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:EVA INT:EVALUACION Psicologica:REPORTE"
    End If
End Sub

Private Sub M1601070000_Click(Index As Integer)
    If Index = 1 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeMantenimiento, RHEstadosTpoPermisosLicencias, "RECURSOS HUMANOS:PERMISOS:APROBACION/RECHAZO"
    ElseIf Index = 2 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeReporte, RHEstadosTpoPermisosLicencias, "RECURSOS HUMANOS:PERMISOS:APROBACION/RECHAZO"
    End If
End Sub

Private Sub M1601070100_Click(Index As Integer)
    If Index = 0 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeRegistro, RHEstadosTpoPermisosLicencias, ""
    ElseIf Index = 1 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeConsulta, RHEstadosTpoPermisosLicencias, ""
    End If
End Sub

Private Sub M1601080000_Click(Index As Integer)
    If Index = 0 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeMantenimiento, RHEstadosTpoVacaciones, "RECURSOS HUMANOS:VACACIONES:MANTENIMIENTO"
    ElseIf Index = 1 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeConsulta, RHEstadosTpoVacaciones, "RECURSOS HUMANOS:VACACIOBES:CONSULTA"
    End If
End Sub

Private Sub M1601090000_Click(Index As Integer)
    If Index = 0 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeMantenimiento, RHEstadosTpoSubsidiado, "RECURSOS HUMANOS:DESCANSOS:MANTENIMIENTO"
    ElseIf Index = 1 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeConsulta, RHEstadosTpoSubsidiado, "RECURSOS HUMANOS:DESCANSOS:CONSULTA"
    End If
End Sub

Private Sub M1601100000_Click(Index As Integer)
    If Index = 0 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeMantenimiento, RHEstadosTpoSuspendido, "RECURSOS HUMANOS:SANCIONES:MANTENIMIENTO"
    ElseIf Index = 1 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeConsulta, RHEstadosTpoSuspendido, "RECURSOS HUMANOS:SANCIONES:CONSULTA"
    End If
End Sub

Private Sub M1601110000_Click(Index As Integer)
    If Index = 1 Then
        frmRHMerDem.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:MERITOS Y DEMERITOS:TABLA:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHMerDem.Ini gTipoOpeConsulta, "RECURSOS HUMANOS:MERITOS Y DEMERITOS:TABLA:CONSULTA"
    End If
End Sub

Private Sub M1601110100_Click(Index As Integer)
    If Index = 0 Then
        frmRHMerDemTabla.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:MERITOS Y DEMERITOS:TABLA:MANTENIMIENTO"
    ElseIf Index = 1 Then
        frmRHMerDemTabla.Ini gTipoOpeReporte, "RECURSOS HUMANOS:MERITOS Y DEMERITOS:TABLA:CONSULTA"
    End If
End Sub

Private Sub M1601120000_Click(Index As Integer)
    If Index = 1 Then
        frmRHEmpleado.Ini gTipoOpeRegistro, RHContratoMantTpoCargo, "RECURSOS HUMANOS:CARGOS LABORALES:REGISTRO"
    End If
End Sub

Private Sub M1601120100_Click(Index As Integer)
    If Index = 0 Then
        frmRHCargos.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:CARGOS LABORALES:TABLA:MANTENIMIENTO"
    ElseIf Index = 1 Then
        frmRHCargos.Ini gTipoOpeConsulta, "RECURSOS HUMANOS:CARGOS LABORALES:TABLA:CONSULTA"
    End If
End Sub

Private Sub M1601130000_Click(Index As Integer)
    If Index = 0 Then
        frmRHEmpleado.Ini gTipoOpeRegistro, RHContratoMantTpoSueldo, "RECURSOS HUMANOS:SUELDO:REGISTRO"
    End If
End Sub

Private Sub M1601140000_Click(Index As Integer)
    If Index = 1 Then
        frmRHEmpleado.Ini gTipoOpeRegistro, RHContratoMantTpoSisPens, "RECURSOS HUMANOS:SISTEMA PENSIONES:REGISTRO"
    End If
End Sub

Private Sub M1601140100_Click(Index As Integer)
    If Index = 0 Then
        frmRHAFP.Ini "RECURSOS HUMANOS:SISTEMA PENSIONES:TABLA:MANTENIMIENTO"
    End If
End Sub

Private Sub M1601150000_Click(Index As Integer)
    If Index = 0 Then
        frmRHInformeSocial.Ini gTipoOpeRegistro, "RECURSOS HUMANOS:INFORME SOCIAL:REGISTRO"
    ElseIf Index = 1 Then
        frmRHInformeSocial.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:INFORME SOCIAL:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHInformeSocial.Ini gTipoOpeConsulta, "RECURSOS HUMANOS:INFORME SOCIAL:CONSULTA"
    ElseIf Index = 3 Then
        frmRHInformeSocial.Ini gTipoOpeReporte, "RECURSOS HUMANOS:INFORME SOCIAL:REPORTE"
    End If
End Sub

Private Sub M1601160000_Click(Index As Integer)
    If Index = 1 Then
        frmRHEmpleado.Ini gTipoOpeRegistro, RHContratoMantTpoAMP, "RECURSOS HUMANOS:ASISTENCIA MEDICA;REGISTRO"
    End If
End Sub

Private Sub M1601160100_Click(Index As Integer)
    If Index = 0 Then
        frmRHAsistMedPrivada.Ini "RECURSOS HUMANOS:ASITENCIA MEDICA:TABLA:MANTENIMIENTO"
    End If
End Sub

Private Sub M1601170100_Click(Index As Integer)
    If Index = 0 Then
        frmRHConceptoMant.Ini gTipoOpeMantenimiento
    ElseIf Index = 1 Then
        frmRHTablasAlias.Show 1
    ElseIf Index = 2 Then
        frmRHConceptoMant.Ini gTipoOpeConsulta
    ElseIf Index = 3 Then
        frmRHConceptoMant.Ini gTipoOpeReporte
    End If
End Sub

Private Sub M1601170200_Click(Index As Integer)
    If Index = 0 Then
        frmRHConceptoAsigPla.Ini gTipoOpeRegistro, ""
    ElseIf Index = 1 Then
        frmRHConceptoAsigPla.Ini gTipoOpeMantenimiento, ""
    ElseIf Index = 2 Then
        frmRHConceptoAsigPla.Ini gTipoOpeConsulta, ""
    ElseIf Index = 3 Then
        frmRHConceptoAsigPla.Ini gTipoOpeReporte, ""
    End If
End Sub

Private Sub M1601180100_Click(Index As Integer)
    If Index = 0 Then
        frmRHConceptoAsigRRHH.Ini gTipoOpeMantenimiento, True
    ElseIf Index = 1 Then
        frmRHConceptoAsigRRHH.Ini gTipoOpeConsulta, True
    ElseIf Index = 2 Then
        frmRHConceptoAsigRRHH.Ini gTipoOpeReporte, True
    End If
End Sub

Private Sub M1601180200_Click(Index As Integer)
    If Index = 0 Then
        frmRHConceptoAsigRRHH.Ini gTipoOpeMantenimiento, False
    ElseIf Index = 1 Then
        frmRHConceptoAsigRRHH.Ini gTipoOpeConsulta, False
    ElseIf Index = 2 Then
        frmRHConceptoAsigRRHH.Ini gTipoOpeReporte, False
    End If
End Sub

Private Sub M1601180300_Click(Index As Integer)
    If Index = 0 Then
        frmRHPlanillaMovExtra.Ini gTipoOpeRegistro, "RECURSOS HUMANOS:PLANILLA DE REMUNERACIONES:EXTRA PLANILLA:REGISTRO"
    ElseIf Index = 1 Then
        frmRHPlanillaMovExtra.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:PLANILLA DE REMUNERACIONES:EXTRA PLANILLA:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHPlanillaMovExtra.Ini gTipoOpeConsulta, "RECURSOS HUMANOS:PLANILLA DE REMUNERACIONES:EXTRA PLANILLA:CONSULTA"
    ElseIf Index = 3 Then
        frmRHPlanillaMovExtra.Ini gTipoOpeReporte, "RECURSOS HUMANOS:PLANILLA DE REMUNERACIONES:EXTRA PLANILLA:REPORTE"
    End If
End Sub

Private Sub M1601190000_Click(Index As Integer)
    Me.Enabled = False
    If Index = 0 Then
        frmRHPlanillas.Ini gTipoProcesoRRHHCalculo, "RECUROS HUMANOS:PROCESOS:CALCULO DE PLANILLAS", Me
    ElseIf Index = 1 Then
        frmRHPlanillas.Ini gTipoProcesoRRHHAbono, "RECUROS HUMANOS:PROCESOS:ABONO DE PLANILLAS", Me
    ElseIf Index = 2 Then
        frmRHCierreMes.Ini "RECUROS HUMANOS:PROCESOS:CIERRE MES"
    ElseIf Index = 3 Then
        frmRHCierreDia.Ini "RECUROS HUMANOS:PROCESOS:CIERRE DIA"
    End If
    Me.Enabled = True
End Sub

Private Sub M1601200000_Click(Index As Integer)
    Me.Enabled = False
    If Index = 0 Then
        frmRRHHRep.Ini "RECURSOS HUMANOS:REPORTES:REPORTES", Me
    ElseIf Index = 1 Then
        frmRRHHRepGen.Ini "RECURSOS HUMANOS:REPORTES:REPORTES GENERALES", Me
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
'    SBBarra.Panels(2).Text = Format(gdFecSis, "dddd - dd - mmmm - yyyy") & Space(3) & Format(Time, "hh:mm AMPM")
'End Sub

Private Sub M1701000000_Click(Index As Integer)
    If Index = 0 Then
        frmLogUsuario.Show 1
    ElseIf Index = 2 Then
        frmLogBieSerMant.Show 1
    ElseIf Index = 4 Then
        frmLogProvMant.Show 1
    ElseIf Index = 13 Then
        frmLogSelConsol.Inicio "3"
    ElseIf Index = 14 Then
        frmLogSelConsol.Inicio "2"
    ElseIf Index = 15 Then
        frmLogSelConsol.Inicio "4"
    ElseIf Index = 16 Then
        frmLogSelConsol.Inicio "5"
    ElseIf Index = 17 Then
        frmLogSelConsol.Inicio "1"
    ElseIf Index = 22 Then
        frmLogAFDeprecia.Show 1
    ElseIf Index = 18 Then
    '   frmLogSelConsol.Inicio "4"
    ElseIf Index = 23 Then
        frmLogOperacionesIngBS.Inicio "58"
    ElseIf Index = 24 Then
        frmTransferencia.Show 1
    End If
End Sub

Private Sub M1701010000_Click(Index As Integer)
    If Index = 0 Then
        frmLogReqInicio.Inicio "1", "1"
    ElseIf Index = 1 Then
        frmLogReqTramite.Inicio "1"
    ElseIf Index = 2 Then
        frmLogReqPrecio.Inicio "1", "1"
    ElseIf Index = 3 Then
        frmLogReqPrecio.Inicio "1", "3"
    ElseIf Index = 4 Then
        frmLogReqPrecio.Inicio "1", "2"
    End If
End Sub

Private Sub M1701020000_Click(Index As Integer)
    If Index = 0 Then
        frmLogReqInicio.Inicio "2", "1"
    ElseIf Index = 1 Then
        frmLogReqTramite.Inicio "2"
    ElseIf Index = 2 Then
        frmLogReqPrecio.Inicio "2", "1"
    ElseIf Index = 3 Then
        frmLogReqPrecio.Inicio "2", "2"
    ElseIf Index = 4 Then
        frmLogReqPrecio.Inicio "2", "3"
    End If
End Sub

Private Sub M1701030000_Click(Index As Integer)
    If Index = 0 Then
        frmLogAdqConsol.Show 1
    ElseIf Index = 1 Then
        frmLogAdqConsul.Show 1
    End If
End Sub


Private Sub M1701040000_Click(Index As Integer)
    If Index = 0 Then
        frmLogSelParaMant.Show 1
    ElseIf Index = 5 Then
        frmLogSelCotPro.Inicio "2", "0"
    End If
End Sub

Private Sub M1701040101_Click(Index As Integer)
    frmLogSelInicio.Inicio IIf(Index = 0, "1", "2"), IIf(Index = 0, False, True)
End Sub

Private Sub M1701040102_Click(Index As Integer)
    frmLogSelInicio.Inicio IIf(Index = 0, "2", "3"), IIf(Index = 0, False, True)
End Sub

Private Sub M1701040103_Click(Index As Integer)
    frmLogSelInicio.Inicio IIf(Index = 0, "3", "4"), IIf(Index = 0, False, True)
End Sub

Private Sub M1701040104_Click(Index As Integer)
    frmLogSelInicio.Inicio IIf(Index = 0, "4", "5"), IIf(Index = 0, False, True)
End Sub

Private Sub M1701040105_Click(Index As Integer)
    frmLogSelInicio.Inicio IIf(Index = 0, "5", "6"), IIf(Index = 0, False, True)
End Sub

Private Sub M1701040106_Click(Index As Integer)
    frmLogSelInicio.Inicio IIf(Index = 0, "6", "7"), IIf(Index = 0, False, True)
End Sub

Private Sub M1701040301_Click(Index As Integer)
    frmLogSelInicio.Inicio IIf(Index = 0, "7", "8"), IIf(Index = 0, False, True)
End Sub

Private Sub M1701040302_Click(Index As Integer)
    frmLogSelInicio.Inicio IIf(Index = 0, "8", "9"), IIf(Index = 0, False, True)
End Sub

Private Sub M1701040303_Click(Index As Integer)
    frmLogSelInicio.Inicio IIf(Index = 0, "9", "10"), IIf(Index = 0, False, True)
End Sub

Private Sub M1701040304_Click(Index As Integer)
    frmLogSelInicio.Inicio IIf(Index = 0, "10", "11"), IIf(Index = 0, False, True)
End Sub

Private Sub M1701040400_Click(Index As Integer)
    If Index = 0 Then
        frmLogSelCotPro.Inicio "1", "1"
    ElseIf Index = 1 Then
        frmLogSelCotPro.Inicio "1", "2"
    ElseIf Index = 2 Then
        frmLogSelCotPro.Inicio "1", "3"
    End If
End Sub

Private Sub M1701050100_Click(Index As Integer)
    If Index = 0 Then
        frmLogOCompra.Inicio False, "501205"
    ElseIf Index = 1 Then
        frmLogOCompra.Inicio False, "502205"
    ElseIf Index = 2 Then
        frmLogOCompra.Inicio False, "501207", "", False
    ElseIf Index = 3 Then
        frmLogOCompra.Inicio False, "502207", "", False
    End If
End Sub

Private Sub M1701050200_Click(Index As Integer)
    If Index = 0 Then
        frmLogOCAtencion.Inicio True, "501206", False, True
    ElseIf Index = 1 Then
        frmLogOCAtencion.Inicio True, "502206", False, True
    ElseIf Index = 2 Then
        frmLogOCAtencion.Inicio False, "501208", False, True
    ElseIf Index = 3 Then
        frmLogOCAtencion.Inicio False, "502208", False, True
    End If
End Sub

Private Sub M1701050300_Click(Index As Integer)
    If Index = 0 Then
        frmLogOCAtencion.Inicio True, "501210", False, False, True
    ElseIf Index = 1 Then
        frmLogOCAtencion.Inicio True, "502210", False, False, True
    ElseIf Index = 2 Then
        frmLogOCAtencion.Inicio False, "501211", False, False, True
    ElseIf Index = 3 Then
        frmLogOCAtencion.Inicio False, "502211", False, False, True
    End If
End Sub

Private Sub M1701050400_Click(Index As Integer)
    If Index = 0 Then
        frmLogOCAtencion.Inicio True, "501210", False, False, True, True
    ElseIf Index = 1 Then
        frmLogOCAtencion.Inicio True, "502210", False, False, True, True
    ElseIf Index = 2 Then
        frmLogOCAtencion.Inicio False, "501211", False, False, True, True
    ElseIf Index = 3 Then
        frmLogOCAtencion.Inicio False, "502211", False, False, True, True
    End If
End Sub

Private Sub M1701060000_Click(Index As Integer)
    If Index = 0 Then
        frmLogOCAtencion.Inicio True, "501221"
    ElseIf Index = 1 Then
        frmLogOCAtencion.Inicio False, "501222"
    ElseIf Index = 2 Then
        frmLogOCAtencion.Inicio True, "502221"
    ElseIf Index = 3 Then
        frmLogOCAtencion.Inicio False, "502222"
    End If
End Sub

Private Sub M1701070000_Click(Index As Integer)
    If Index = 0 Then
        frmLogOperacionesIngBS.Inicio "59"
    ElseIf Index = 1 Then
        frmLogAlmInven.Show 1
    ElseIf Index = 2 Then
        frmLogKardex.Show 1
    ElseIf Index = 3 Then
        frmLogCalculaSaldos.Show 1
    ElseIf Index = 4 Then
        'frmMantPermisos.Show 1
    ElseIf Index = 5 Then
        frmLogMantSaldos.Show 1
    ElseIf Index = 6 Then
        frmLogCalculaAsientos.Show 1
    End If
End Sub

Private Sub M1701080000_Click(Index As Integer)
    '1 - REGISTRO ;     2 - DISTRIBUCION;     3 - GARANTIA
    frmLogSerCon.Inicio (Index + 1)
End Sub

Private Sub M1801000000_Click(Index As Integer)
    If Index = 0 Then
        frmPreMantenimiento.Show 1
    ElseIf Index = 1 Then
        frmPreRubros.Show 1
    ElseIf Index = 2 Then
        frmPlaPresu.Show 1
    ElseIf Index = 3 Then
        frmPlaEjecu.Show 1
    End If
End Sub

Private Sub M1801010000_Click(Index As Integer)
    If Index = 0 Then
        frmLogOCAtencion.Inicio True, "501221", True
    Else
        frmLogOCAtencion.Inicio True, "502221", True
    End If
End Sub

Private Sub M1801020000_Click(Index As Integer)
    If Index = 0 Then
        frmLogOCAtencion.Inicio False, "501222", True
    Else
        frmLogOCAtencion.Inicio False, "502222", True
    End If
End Sub

Private Sub MDIForm_Load()
    'Dim Cont As Control
    'On Error Resume Next
    'For Each Cont In Controls ' Itera por cada elemento.
    '    If Left(Cont.Name, 1) = "M" Then
    '        If Left(Cont.Name, 3) <> "M09" And Left(Cont.Name, 3) <> "M18" And Left(Cont.Name, 3) <> "M17" And Left(Cont.Name, 3) <> "M01" And Left(Cont.Name, 3) <> "M16" Then
    '        'If Left(Cont.Name, 3) <> "M09" And Left(Cont.Name, 3) <> "M18" And Left(Cont.Name, 3) <> "M17" And Left(Cont.Name, 3) <> "M01" Then
    '          Cont.Visible = False
    '        End If
    '    End If
    'Next
End Sub

