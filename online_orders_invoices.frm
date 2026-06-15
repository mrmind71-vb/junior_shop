VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form orders_online_invoices 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«’‰«› ÿ·»Ì« "
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   20400
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   8970
   ScaleWidth      =   20400
   Visible         =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame fmReplace 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   1485
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   1080
      Width           =   3525
      Begin Threed.SSCommand xSales_replace 
         Height          =   330
         Left            =   90
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   225
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   582
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TagVariant      =   "«Œ «— ‰Ê⁄ «·ÿ·»Ì…"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "›« Ê—… «” »œ«·"
         Height          =   270
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   270
         Width           =   1080
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   1080
      Width           =   2445
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   225
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«· «—ÌŒ"
         Height          =   285
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   270
         Width           =   690
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1185
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   1185
      ScaleWidth      =   20400
      TabIndex        =   21
      Top             =   7410
      Width           =   20400
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1050
         Left            =   12870
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   -45
         Width           =   7485
         Begin VB.Label xrate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4320
            TabIndex        =   38
            Top             =   585
            Width           =   600
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "≈Ã„«·Ì «·ﬂ„Ì…"
            Height          =   285
            Left            =   6075
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   270
            Width           =   1185
         End
         Begin VB.Label xtotalQuant 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4320
            TabIndex        =   34
            Top             =   225
            Width           =   1635
         End
         Begin VB.Label xtotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   90
            TabIndex        =   33
            Top             =   585
            Width           =   1635
         End
         Begin VB.Label xdiscount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4950
            TabIndex        =   32
            Top             =   585
            Width           =   1005
         End
         Begin VB.Label xtotal_item 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   90
            TabIndex        =   31
            Top             =   225
            Width           =   1635
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "«·Œ’„"
            Height          =   285
            Left            =   6075
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   630
            Width           =   735
         End
         Begin VB.Label label50 
            BackColor       =   &H00FFFFFF&
            Caption         =   "’«›Ì «·ﬁÌ„…"
            Height          =   285
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   630
            Width           =   1185
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "≈Ã„«·Ì «·«’‰«›"
            Height          =   285
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   270
            Width           =   1320
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   0
         Width           =   3570
         Begin Threed.SSCommand cmdFirst 
            Height          =   510
            Left            =   2700
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   135
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   900
            _Version        =   196610
            BackColor       =   16777215
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "online_orders_invoices.frx":0000
            Caption         =   "√Ê·"
            ButtonStyle     =   3
            PictureAlignment=   10
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "online_orders_invoices.frx":21A7
         End
         Begin Threed.SSCommand cmdPrevious 
            Height          =   510
            Left            =   1800
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   135
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   900
            _Version        =   196610
            BackColor       =   16777215
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "online_orders_invoices.frx":41EE
            Caption         =   "”«»ﬁ"
            ButtonStyle     =   3
            PictureAlignment=   10
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "online_orders_invoices.frx":62D9
         End
         Begin Threed.SSCommand cmdNext 
            Height          =   510
            Left            =   945
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   135
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   900
            _Version        =   196610
            BackColor       =   16777215
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "online_orders_invoices.frx":82D3
            Caption         =   "·«Õﬁ"
            ButtonStyle     =   3
            PictureAlignment=   9
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "online_orders_invoices.frx":A3E4
         End
         Begin Threed.SSCommand cmdLast 
            Height          =   510
            Left            =   0
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   135
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   900
            _Version        =   196610
            BackColor       =   16777215
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "online_orders_invoices.frx":C3DE
            Caption         =   "√ŒÌ—"
            ButtonStyle     =   3
            PictureAlignment=   9
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "online_orders_invoices.frx":E602
         End
      End
      Begin Threed.SSCommand cmdSend 
         Height          =   510
         Left            =   3690
         TabIndex        =   42
         Top             =   135
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "online_orders_invoices.frx":106D3
         Caption         =   " —ÕÌ· ·›« Ê—… „»Ì⁄« "
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_invoices.frx":12DF6
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   9675
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   720
      Width           =   10590
      Begin VB.TextBox xNotes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   585
         Width           =   5910
      End
      Begin VB.TextBox xShip_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   7740
         RightToLeft     =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   585
         Width           =   1455
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   7740
         RightToLeft     =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   225
         Width           =   1455
      End
      Begin Threed.SSCommand cmdType 
         Height          =   330
         Left            =   3330
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   225
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   582
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "‰Ê⁄ «·ÿ·»Ì…"
         TagVariant      =   "«Œ «— ‰Ê⁄ «·ÿ·»Ì…"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin VB.Label xtype 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   4410
         TabIndex        =   39
         Tag             =   "1"
         Top             =   225
         Width           =   1590
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "‰Ê⁄ «·ÿ·»Ì…"
         Height          =   270
         Left            =   6075
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„·ÕÊŸ…"
         Height          =   285
         Left            =   6075
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   630
         Width           =   690
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "»Ê·Ì’… «·‘Õ‰"
         Height          =   285
         Left            =   9270
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   630
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«· «—ÌŒ"
         Height          =   330
         Left            =   9315
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   225
         Width           =   690
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   15120
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   45
      Width           =   5145
      Begin Threed.SSCommand cmdNewInv 
         Height          =   510
         Left            =   2655
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "online_orders_invoices.frx":150EC
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_invoices.frx":170F4
      End
      Begin Threed.SSCommand cmddel 
         Height          =   510
         Left            =   1350
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   135
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "online_orders_invoices.frx":190AB
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_invoices.frx":1B847
      End
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   3915
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "online_orders_invoices.frx":1DCDB
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_invoices.frx":200A6
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   135
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "online_orders_invoices.frx":2214F
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame FRAME_CUR 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Index           =   4
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   630
      Width           =   1545
      Begin Threed.SSCommand cmdSave 
         Height          =   465
         Left            =   45
         TabIndex        =   11
         Top             =   135
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   820
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "online_orders_invoices.frx":24472
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_invoices.frx":26D97
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   465
         Left            =   45
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   630
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   820
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "online_orders_invoices.frx":295EB
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_invoices.frx":2B74B
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   465
      Left            =   -1845
      Top             =   90
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   820
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   8595
      Width           =   20400
      _ExtentX        =   35983
      _ExtentY        =   661
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Top             =   45
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   476
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   1
         Left            =   3240
         TabIndex        =   8
         Top             =   45
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   2
         Left            =   7110
         TabIndex        =   9
         Top             =   45
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   5550
      Left            =   135
      TabIndex        =   4
      Top             =   1800
      Width           =   20175
      _cx             =   35586
      _cy             =   9790
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   12648384
      ForeColorFixed  =   0
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483630
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame fmSend 
      BackColor       =   &H00FFFFFF&
      Height          =   1050
      Left            =   1395
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   720
      Width           =   3615
      Begin VB.Label lblReturn 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "›« Ê—… «·„— Ã⁄"
         Height          =   270
         Left            =   2115
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   630
         Width           =   1110
      End
      Begin VB.Label xSales_ret 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   90
         TabIndex        =   47
         Tag             =   "1"
         Top             =   585
         Width           =   1905
      End
      Begin VB.Label xdoc_no_sales 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   90
         TabIndex        =   46
         Tag             =   "1"
         Top             =   225
         Width           =   1905
      End
      Begin VB.Label lblSend 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·›« Ê—… «·„—Õ·…"
         Height          =   270
         Left            =   2115
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   270
         Width           =   1215
      End
   End
   Begin VB.Label xdoc_no 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   0
      TabIndex        =   43
      Tag             =   "1"
      Top             =   0
      Visible         =   0   'False
      Width           =   2040
   End
End
Attribute VB_Name = "orders_online_invoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sOrder_No As String
Public sSales_doc As String
Public sStore As String
Public myForm As Form
Public bEdit As Boolean
Dim temptable As New ADODB.Recordset, cMsgDisplay As String, nDayDiscout As Double, cDescDiscout As String
Dim bAct As Boolean
Dim bUpdated As Boolean
Dim cFilter As String
Dim bIg As Boolean
Dim bStopCell As Boolean
Dim cFilterLook As String
Dim nRound As Long
Dim CardTable As ADODB.Recordset
Dim oSearchItem As New Search_abd, oSearchDoc   As New Search_abd
Dim oSearchDocRet As New Search_abd
Dim bEditRecord As Boolean
Dim oSalesRefund As New sales_refundfrm
Dim oSearchType As New search_empty
Dim formMode
Const LoadMode = 0, DefineMode = 1
Dim nUser As Long
Private Enum enUser
MainUser = 1
branch = 0
End Enum
Sub ItemsLookup()
ItemsLookupAll Me, oSearchItem
End Sub
Function myreplace(Optional Row As Long = -1, Optional bOffer As Boolean, Optional bReplace As Boolean) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(aInsert, "[NOTES]", addstring(xNotes.text))
aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(xdiscount.Caption))
aInsert = AddFlag(aInsert, "[TYPE]", addvalue(xtype.Tag))
aInsert = AddFlag(aInsert, "[SALES_REPLACE]", addstring(xSales_replace.Caption))
aInsert = AddFlag(aInsert, "[DATE]", addDate(xDate.text))
aInsert = AddFlag(aInsert, "[ORDER_NO]", addstring(sOrder_No))
aInsert = AddFlag(aInsert, "[DATE1]", addDate(xDate1.text))

On Error GoTo myerror
Dim con As New ADODB.Connection
If openCon(con) <> "ok" Then Exit Function
con.BeginTrans
If xdoc_no.Tag = DefineMode Then
    xdoc_no.Caption = Newflag("FILE6_90BH", "DOC_NO", con)
    aInsert = AddFlag(aInsert, "DOC_NO", addstring(xdoc_no.Caption))
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(cUserName))
    aInsert = AddFlag(aInsert, "[USER_IP]", addstring(GetComputerName))
    con.Execute addInsert(aInsert, "FILE6_90BH")
Else
    con.Execute addUpdate(aInsert, "FILE6_90BH", "DOC_NO = " & addstring(xdoc_no.Caption))
End If

If bOffer Then
'    myreplaceGrdOffer
'    UpdateDiscount xdoc_no.text, con
Else
    myreplaceGrd Row, con
End If
con.CommitTrans
myreplace = True
bUpdated = True
Finally:
closeCon con
Exit Function
myerror:
MsgBox Err.Description
If TransCount(con) > 0 Then con.RollbackTrans
Err.Clear
GoTo Finally
End Function
Sub myProc(Optional sControl As String = "")
On Error GoTo myerror
Dim sDoc As String
If ActiveControl.Name = grid1.Name Then
    Dim bNew As Boolean
    bNew = grid1.Row = grid1.Rows - 1
    grid1.TextMatrix(grid1.Row, 1) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    If Not bNew Then
        CellPos 13, grid1.Row, grid1.col
    End If
    GrdDesc grid1.TextMatrix(grid1.Row, 1), grid1.Row
    Grid1_AfterEdit grid1.Row, grid1.col
ElseIf ActiveControl.Name = xSales_replace.Name Then
    grid1.Row = grid1.Rows - 1
    grid1.TextMatrix(grid1.Row, 1) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    If Not bNew Then
        CellPos 13, grid1.Row, grid1.col
    End If
    GrdDesc grid1.TextMatrix(grid1.Row, 1), grid1.Row
    Grid1_AfterEdit grid1.Row, grid1.col
ElseIf ActiveControl.Name = cmdInform.Name Then
    xdoc_no.Caption = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    Unload oSearchDoc
    myUndo
ElseIf sControl = "inv_replace" Then
    sDoc = oSearchDocRet.grid1.TextMatrix(oSearchDocRet.grid1.Row, 0)
    Unload oSearchDocRet
    xSales_replace.Caption = sDoc
ElseIf sControl = "inv_ret" Then
    sDoc = oSearchDocRet.grid1.TextMatrix(oSearchDocRet.grid1.Row, 0)
    Unload oSearchDocRet
    addRefundOrder sDoc
ElseIf ActiveControl.Name = cmdType.Name Then
    xSales_replace.Caption = ""
    
    fmReplace.Visible = False
    fmSend.Visible = False
        
    If oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 0) = "2" Then
        Unload oSearchType
        addRefund
    ElseIf oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 0) = "3" Or oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 0) = "4" Then
        addReplace
    Else
        xtype.Tag = oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 0)
        xtype.Caption = oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 1)
        Unload oSearchType
        Handlecontrols xdoc_no.Tag
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Function addRefund() As Integer
Dim aValues As Variant
aValues = rsValues("SELECT COUNT(*) AS COUNT_OF,MIN(DOC_NO) AS DOC_NO FROM FILE6_20H WHERE IS_RETURN = 0 AND  ONLINE_DOC = " & MyParn(sOrder_No))
If IsEmpty(aValues) Then
    MsgBox "·«  ÊÃœ ›Ê« Ì— „— Ã⁄"
    Exit Function
End If

If retFlag(aValues, "COUNT_OF") = 1 Then
    addRefundOrder retFlag(aValues, "DOC_NO") & ""
Else
    invLook "inv_ret"
End If
End Function
Private Function addReplace() As Integer
xtype.Tag = oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 0)
xtype.Caption = oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 1)
Unload oSearchType

Dim aValues As Variant
aValues = rsValues("SELECT COUNT(*) AS COUNT_OF,MIN(DOC_NO) AS DOC_NO FROM FILE6_20H WHERE IS_RETURN = 0 AND  ONLINE_DOC = " & MyParn(sOrder_No))
If IsEmpty(aValues) Then
    MsgBox "·«  ÊÃœ ›Ê« Ì— ··«” »œ«·"
    Exit Function
End If

If retFlag(aValues, "COUNT_OF") = 1 Then
    xSales_replace.Caption = retFlag(aValues, "DOC_NO") & ""
    Handlecontrols xdoc_no.Tag
Else
    invLook "inv_repalce"
End If

fmReplace.Visible = True
End Function
Private Sub invLook(Optional sControl As String)
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(4, 1)
Set Generalarray(0) = Me

Generalarray(1) = "SELECT TOP 1000 " & _
                  " FILE6_20H.DOC_NO ," & _
                  " FILE6_20H.DOC_NO2," & _
                  " FORMAT(FILE6_20H.date,'yyyy/M/d')," & _
                  " FILE6_20H.TOTAL_ITEM - FILE6_20H.DISCOUNT," & _
                  " ONLINE_TYPE_CODES.DESCA" & _
                  " FROM FILE6_20H " & _
                  " LEFT JOIN ONLINE_TYPE_CODES ON FILE6_20H.INV_TYPE_ONLINE = ONLINE_TYPE_CODES.CODE" & _
                  " WHERE FILE6_20H.ONLINE_DOC = " & MyParn(sOrder_No) & _
                  " AND FILE6_20H.IS_RETURN = 0"

Generalarray(2) = "Order by DATE DESC,doc_no2 DESC "
Generalarray(3) = 6000
Generalarray(5) = False

listarray(0, 0) = "—ﬁ„ «·„” ‰œ"
listarray(0, 1) = "(FILE6_20H.DOC_NO = 'cFilter' OR FILE6_20H.DOC_NO2 = 'cFilter')"

listarray(1, 0) = "«· «—ÌŒ"
listarray(1, 1) = "(##FILE6_20H.DATE##)"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(1, 1) = 1200

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "«·≈Ã„«·Ì"
GrdArray(3, 1) = 1300

GrdArray(4, 0) = "‰Ê⁄ «·„” ‰œ"
GrdArray(4, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDocRet.Caption = "«” ⁄·«„"
oSearchDocRet.sControl = sControl
oSearchDocRet.Show 1
End Sub
Private Function addRefundOrder(pDoc_no) As Boolean
Dim loctable As New ADODB.Recordset
Dim cString As String
cString = "select TOP 1 FILE6_20H.BRANCH," & _
          " FILE6_20H.PRINTED," & _
          " FILE6_20H.ISINVOICE," & _
          " FILE6_20H.TOTAL_ITEM," & _
          " FILE6_20H.DATE, " & _
          " FILE6_20H.ISRET" & _
          " FROM FILE6_20H " & _
          " WHERE FILE6_20H.DOC_NO = " & MyParn(pDoc_no)

Set loctable = myRs(cString)
If loctable.EOF Then
    MsgBox "›« Ê—… €Ì— „”Ã·…"
    Exit Function
End If

If loctable!branch & "" = "00" Then
    MsgBox "›« Ê—… ›—⁄ —∆Ì”Ì"
    Exit Function
End If

If Not loctable!PRINTED Then
    MsgBox "›« Ê—… „› ÊÕ…"
    Exit Function
End If

If loctable!ISINVOICE Then
    MsgBox "›« Ê—… „Ã„⁄…"
    Exit Function
End If

If loctable!TOTAL_ITEM <= 0 Then
    If loctable!TOTAL_ITEM = 0 Then
        MsgBox "›« Ê—… »œÊ‰ ﬁÌ„…"
    End If
    If loctable!TOTAL_ITEM < 0 Then
        MsgBox "›« Ê—… „— Ã⁄"
    End If
    Exit Function
End If

Set oSalesRefund.myForm = Me
oSalesRefund.sDoc_no = pDoc_no
oSalesRefund.Show 1
End Function
Private Sub CmdDel_Click()
If MsgBox("Õ–›", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub

Dim con As New ADODB.Connection
openCon con
con.BeginTrans
con.Execute "Delete  From FILE6_90B where Doc_No = " & xdoc_no.Caption
con.Execute "Delete  From FILE6_90BH where Doc_No = " & xdoc_no.Caption
con.CommitTrans

If Not openCardTable(tbMode.tbPrevious, xdoc_no.Caption) Then
    If Not openCardTable(tbMode.tbFirst) Then
        myDefine
    End If
End If
bUpdated = True
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CmdExit_Click()
    sDoc_no = ""
    Unload Me
End Sub
Private Sub CmdInform_Click()
CardLookup
End Sub
Private Sub CmdNewInv_Click()
myDefine
On Error Resume Next
grid1.Select 1, 1
Err.Clear
End Sub
Private Sub cmdSave_Click()
If grid1.Rows = 2 Then Exit Sub
If Not myValid Then Exit Sub
If myreplace Then
    Inform " „ «·Õ›Ÿ"
    If Not openCardTable(tbMode.tbFind, xdoc_no.Caption) Then
        If Not openCardTable Then
            myDefine
        End If
    End If
End If
End Sub
Private Sub cmdsend_Click()
sdoc_no_new = myReplaceOnline
If sDoc_no_no <> "" Then
    MsgBox " „ «· —ÕÌ· «·Ì «·›« Ê—… —ﬁ„ : " & sdoc_no_new
End If

If Not openCardTable(tbMode.tbFind, xdoc_no.Caption) Then
    If Not openCardTable Then
        myDefine
    End If
End If
End Sub
Private Sub cmdType_Click()
OnlineTypeLookup Me, oSearchType, IIf(grid1.Rows > 2, "CODE <> 1", "")
End Sub
Private Sub CmdUndo_Click()
myUndo
End Sub
Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    If xdoc_no.Tag = LoadMode Then
        On Error Resume Next
        grid1.SetFocus
        CellPos 13, grid1.Rows - 2, grid1.Cols - 1
    Else
        On Error Resume Next
        cmdType.SetFocus
    End If
End If
Err.Clear
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then KeyAscii = 0
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        SendKeys "{TAB}"
        KeyCode = 0
    End If
End If
End Sub
Private Sub Form_Load()
'On Error GoTo myerror
'openCon con


Me.Caption = "„»»⁄«  " & myFormat_p(dSalesDate)
nRound = 0
HandleInit

'Set xType.RowSource = myRs("SELECT * FROM ONLINE_TYPE_CODES ORDER BY CODE")
'xType.BoundColumn = "CODE"
'xType.ListField = "DESCA"

cDefClient = "0000"

If Not openCardTable Then myDefine
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Form_Resize()
Dim nHeight As Long
nHeight = Me.Height - (grid1.Top + 4100)
If nHeight > 3000 Then
    grid1.Height = nHeight
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
Set orders_online_invoices = Nothing
If bUpdated And Not myForm Is Nothing Then
    myForm.myLoadCount
End If
Err.Clear
End Sub
Public Sub Grid1_AfterEdit(ByVal Row As Long, ByVal col As Long)
'On Error GoTo myerror
Dim bNew As Boolean
With grid1
If Not myValid(True) Then
    On Error Resume Next
    grid1.SetFocus
    Err.Clear
    myLoadGrd
    If Row < grid1.Rows - 1 Then
        grid1.Select Row, col
    Else
        CellPos 13, grid1.Rows - 2, grid1.Cols - 1
    End If
    Exit Sub
End If


If Not validRow(Row) Then
    CalcTotals
    Exit Sub
End If

If Row = grid1.Rows - 1 Then
    myAddItem
    bNew = True
End If

CalcTotals

If myreplace(Row) Then
    If xdoc_no.Tag = DefineMode Then
        openCardTable tbMode.tbFind, xdoc_no.Caption
    ElseIf grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
        myLoadGrd
    End If
    If bNew Then
        grid1.Select grid1.Rows - 1, 1
    End If
End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
grid1_EnterCell
End Sub
Private Sub grid1_EnterCell()
With grid1
    If sBranch <> "00" Then
        .Editable = flexEDNone
    ElseIf Not bEditRecord Then
        .Editable = flexEDNone
    ElseIf xtype.Tag = "" Then
        .Editable = flexEDNone
    ElseIf .ValueMatrix(grid1.Row, 19) > 0 Then
        .Editable = flexEDNone
    ElseIf (.col = 1 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "") Or (.col = 10) Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Private Sub grid1_GotFocus()
grid1_EnterCell
Err.Clear
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    'If Col = 12 And grid1.TextMatrix(Row, Col) = "" Then Exit Sub
    CellPos KeyCode, Row, col
End If
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
Dim sBarCode As String
Dim nSal As Double
Dim nRet As Double
Dim cPassword As String, aUser As Variant
With grid1
    If col = 10 Then
        If Not IsNumeric(.EditText) Then
            Cancel = True
            Exit Sub
        ElseIf Val(.EditText) = 0 Then
            MsgBox "€»— „”„ÊÕ »ﬂ„Ì… ’›—Ì…"
            Cancel = True
        ElseIf Val(.EditText) < 0 Then
            If Val(xtype.Tag) = 1 Or Val(xtype.text) = 3 Then
                MsgBox "ﬂ„Ì… „— Ã⁄ ›Ï " & xtype.Caption
                Cancel = True
                Exit Sub
            End If
'            For i = 1 To .Rows - 1
'                If i <> Row Then
'                    If .ValueMatrix(i, 10) > 0 Then
'                        MsgBox "ﬂ„Ì… „— Ã⁄ ›Ï ›« Ê—… »ﬂ„Ì«  „ÊÃ»…"
'                        Cancel = True
'                        Exit Sub
'                    End If
'                End If
'            Next
        ElseIf Val(.EditText) > 0 Then
            If Val(xtype.Tag) = 2 Or Val(xtype.text) = 3 Then
                MsgBox "ﬂ„Ì… „ÊÃ»… ›Ï  " & xtype.Caption
                Cancel = True
                Exit Sub
            End If

'            For i = 1 To .Rows - 1
'                If i <> Row Then
'                    If .ValueMatrix(i, 10) < 0 Then
'                        MsgBox "ﬂ„Ì… „ÊÃ»… ›Ï ›« Ê—… „— Ã⁄"
'                        Cancel = True
'                        Exit Sub
'                    End If
'                End If
'            Next
        ElseIf Val(.EditText) > 0 Then
            nBalance = rsBalance(grid1.TextMatrix(grid1.Row, 1), cBranchStore, xDate.text, grid1.TextMatrix(grid1.Row, grid1.Cols - 1))
            If nBalance < Val(.EditText) Then
                Inform "«·—’Ìœ ·« Ì”„Õ"
                Exit Sub
            End If
        End If
    ElseIf col = 1 Then
        If Not GrdDesc(.EditText, Row) Then
            Cancel = True
        Else
            If .EditText <> grid1.TextMatrix(Row, 1) Then
                .EditText = grid1.TextMatrix(Row, 1)
            End If
        End If
    End If
End With
End Sub
Private Function myValid(Optional bIgMsg As Boolean = False, Optional bIgMan As Boolean = False) As Boolean
If Not IsDate(xDate.text) Then
    If Not bIgMsg Then MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If xtype.Tag = "" Then
    MsgBox "·„ Ì „  ÕœÌœ ‰Ê⁄ «·ÿ·»Ì…"
    Exit Function
End If

If Trim(xShip_No.text) <> "" Then
    MsgBox "—ﬁ„ »Ê·Ì’… «·‘Õ‰ „ÿ·Ê»"
    Exit Function
End If

'If IsDate(xDate1.text) And xtype.Tag = "3" And Trim(xShip_No.text) = "" Then
'    MsgBox " «—ÌŒ «” ·«„ ›Ï „” ‰œ «” »œ«· »œÊ‰ —ﬁ„ »Ê·Ì’… ‘Õ‰"
'    Exit Function
'End If

If Not validQuant Then
    If Not bIgMsg Then MsgBox "ﬂ„Ì… „ÊÃ»… Êﬂ„Ì… ”«·»… ›Ï ‰›” «·›« Ê—…"
    Exit Function
End If
myValid = True
End Function
Private Sub myload()
xdoc_no.Caption = CardTable!DOC_NO
xDate.text = myFormat_p(CardTable!Date)
xDate1.text = myFormat_p(CardTable!Date1)
xtype.Tag = CardTable!Type
xtype.Caption = CardTable!TYPE_dESCA
xdiscount.Caption = CardTable!discount & ""
xNotes.text = CardTable!NOTES & ""
xdoc_no_sales.Caption = CardTable!doc_no_sales & ""
xSales_replace.Caption = CardTable!sales_replace & ""
xSales_ret.Caption = CardTable!SALES_RET & ""
'xtime.Caption = Format(CardTable!Time, "hh:nn")
panel1(1).Caption = CardTable!user_ip & ""

myLoadGrd

Handlecontrols LoadMode

CellPos 13, grid1.Rows - 2, grid1.Cols - 1
Exit Sub
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub myDefine()
xdoc_no.Caption = ""
xDate.text = myFormat_p(Date)
'xType.BoundText = ""
xdoc_no_sales.Caption = ""
xSales_replace.Caption = ""
xSales_ret.Caption = ""

xtype.Caption = ""
xtype.Tag = ""
xNotes.text = ""
xtotal_item.Caption = ""
xdiscount.Caption = ""
xtotal.Caption = ""

grid1.Rows = 1
myAddItem

Handlecontrols DefineMode
Fixgrd
End Sub
Private Sub Handlecontrols(nMode)
bEditRecord = bEdit
bEditRecord = bEditRecord And cBranch = "00"
bEditRecord = bEditRecord And xtype.Caption <> ""
bEditRecord = bEditRecord And xdoc_no_sales.Caption = ""

'cmdSave.Enabled = bEditRecord Or (xtype.Tag = "3" And (cBranch <> "00" And xDate2.text = ""))
cmdSave.Enabled = bEditRecord Or (xtype.Tag > "3" And (cBranch <> "00" And xDate1.text = ""))
cmdNewInv.Enabled = bEdit And cBranch = "00"
cmddel.Enabled = bEditRecord And nMode = LoadMode And cBranch = "00" And (Trim(xDate1.text) = "" Or xtype.Tag <> "3")

cmdType.Enabled = nMode = DefineMode And cBranch = "00"
cmdSend.Enabled = nMode = LoadMode And xdoc_no_sales.Caption = "" And xtype.Tag <> "3" And cBranch <> "00"

xDate1.Enabled = (cBranch <> "00") Or (cBranch = "00" And IsDate(xDate1.text))

Dim nRecord As Long, nRecords As Long
retRecords xdoc_no.Caption, nRecords, nRecord

cmdNext.Enabled = nRecord < nRecords And nRecords <> 0 And nMode = LoadMode
cmdPrevious.Enabled = nRecord <> 1 And nRecords <> 0 And nMode = LoadMode
cmdLast.Enabled = nRecord < nRecords And nRecords <> 0 And nMode = LoadMode
cmdFirst.Enabled = nRecord <> 1 And nRecords <> 0 And nMode = LoadMode

If nMode = LoadMode Then
    panel1(0).Caption = "”Ã· " & nRecord & " „‰ " & nRecords
Else
    panel1(0).Caption = "”Ã· ÃœÌœ (" & (nRecords + 1) & ")"
End If

fmReplace.Visible = nMode = LoadMode And xtype.Tag = "3"
fmSend.Visible = nMode = LoadMode And xtype.Tag <> "3"

'lblReplace.Visible = nMode = LoadMode And xSales_replace.Caption <> ""
'xSales_replace.Visible = nMode = LoadMode And xSales_replace.Caption = ""

xdoc_no.Tag = nMode
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If Not bEditRecord Then
    Exit Sub
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.col
ElseIf KeyCode = 112 And grid1.Editable = flexEDKbdMouse Then
    ItemsLook Me, oSearchItem
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Rows > 3 And (grid1.TextMatrix(grid1.Row, 18) = "" Or grid1.ValueMatrix(grid1.Row, 19) = 0) Then
    RemoveItem (grid1.Row)
    
    grid1.Select grid1.Rows - 1, 1
    grid1.ShowCell grid1.Rows - 1, 1
    MakeSerial grid1.Row
        
    CalcTotals
End If
End Sub
Private Function GrdDesc(pItem As String, ByRef Row As Long, Optional ByRef sMsg As String) As Boolean
With grid1
If Trim(pItem) = "" Or Len(pItem) < 3 Or Len(pItem) > 30 Then
    sMsg = "ﬂÊœ €Ì— ’«·Õ"
    Exit Function
End If

Dim con As New ADODB.Connection
openCon con

Dim cm As New ADODB.command
Set cm = cmd("dbo.sp_item", con, adStoredProc, AddFlag(Empty, "BARCODE_FIND", pItem))
cm.Execute

If IsNull(cm.Parameters("@ITEM").Value) Then
    sMsg = "ﬂÊœ €Ì— ’«·Õ"
    GoTo Finally
End If

Dim nBalance As Double
If Not cm.Parameters("@ISNOITEM").Value Then
    If cm.Parameters("@PRICE").Value = 0 Then
        MsgBox "”⁄— «·»Ì⁄ ’›— "
        GoTo Finally
    End If
    
'    If fnBalance(cm.Parameters("@ITEM").Value, con, sStore, xDate.text) <= 0 Then
'        Inform "·« ÌÊÃœ —’Ìœ «·’‰›"
'    End If
End If

Dim cmDiscount As New ADODB.command
aPrm = AddFlag(Empty, "ITEM", cm.Parameters("@ITEM").Value)
aPrm = AddFlag(aPrm, "DATE", myFormat_sp(xDate.text))
Set cmDiscount = cmd("dbo.sp_offer_price", con, adStoredProc, aPrm)
cmDiscount.Execute

grid1.TextMatrix(Row, 1) = cm.Parameters("@ITEM").Value
grid1.TextMatrix(Row, 2) = cm.Parameters("@MOSM").Value
grid1.TextMatrix(Row, 3) = cm.Parameters("@FACT").Value
grid1.TextMatrix(Row, 4) = cm.Parameters("@SUPP").Value & ""
grid1.TextMatrix(Row, 5) = cm.Parameters("@MODELFACT0").Value
grid1.TextMatrix(Row, 6) = cm.Parameters("@DESCA").Value
grid1.TextMatrix(Row, 7) = cm.Parameters("@SCAL").Value
grid1.TextMatrix(Row, 8) = cm.Parameters("@COLOR").Value
grid1.TextMatrix(Row, 9) = cm.Parameters("@PRICE").Value
'If grid1.Rows > 2 Then
'    grid1.TextMatrix(Row, 10) = IIf(grid1.ValueMatrix(1, 10) < 0, -1, 1)
'Else
'    grid1.TextMatrix(Row, 10) = "1"
'End If

grid1.TextMatrix(Row, 10) = IIf(xtype.Tag = 1 Or xtype.Tag = 3, 1, -1)
grid1.TextMatrix(i, 14) = fnBalance(grid1.TextMatrix(Row, 1), con, sStore)
If IsNull(cmDiscount.Parameters("@ID").Value) Then
    grid1.TextMatrix(Row, 11) = cm.Parameters("@PRICE").Value
End If
grid1.TextMatrix(Row, 17) = Val(cm.Parameters("@COSTITEM").Value & "")
If Not IsNull(cmDiscount.Parameters("@ID").Value) Then
    grid1.TextMatrix(Row, 11) = cmDiscount.Parameters("@PRICE").Value
    grid1.TextMatrix(Row, 18) = cmDiscount.Parameters("@DOC_NO").Value
    grid1.TextMatrix(Row, 19) = 0
    grid1.TextMatrix(Row, 20) = 0
End If
GrdDesc = True
Finally:
closeCon con
End With
End Function
Private Function CalcTotals(Optional nMode As Integer = 0)
Dim nTotalQuant As Double, nTotalQuant2 As Double, nTotalItem As Double, nTotalItemNoDiscount As Double, nTotalDiscount As Double, nTotalMin As Double
Dim nDisountRow As Double
With grid1
Dim i As Long
For i = 1 To grid1.Rows - 2
    nDisountRow = (.ValueMatrix(i, 9) * .TextMatrix(i, 10)) - (.ValueMatrix(i, 11)) * .ValueMatrix(i, 10)
    .TextMatrix(i, 13) = .ValueMatrix(i, 10) * .ValueMatrix(i, 9)
    
    If nDisountRow <> 0 And .ValueMatrix(i, 13) <> 0 Then
        .TextMatrix(i, 12) = Format(Round(nDisountRow / .ValueMatrix(i, 13), 2), "0%")
    Else
        .TextMatrix(i, 12) = ""
    End If
        
    
    nTotalQuant = nTotalQuant + grid1.ValueMatrix(i, 10)
    nTotalItem = nTotalItem + .ValueMatrix(i, 10) * .ValueMatrix(i, 11)
    
    nTotalItemNoDiscount = nTotalItemNoDiscount + (.ValueMatrix(i, 9) * .ValueMatrix(i, 10))
    nTotalItemDiscount = nTotalItemDiscount + nDisountRow
Next

xtotalQuant.Caption = Myvalue(nTotalQuant)
'xTotal_itemNoDiscount.Caption = Myvalue(nTotalItemNoDiscount)
'xTotal_itemDiscount.Caption = Myvalue(nTotalItemDiscount)
xtotal_item.Caption = Myvalue(nTotalItem)
'xTotal_item_offer.Caption = Myvalue(nTotalItem - Val(xDiscount_offer.Caption))
'If nTotalItemNoDiscount <> 0 And nTotalItemDiscount Then
'    xTotal_itemDiscount_rate.Caption = Format(Round(nTotalItemDiscount / nTotalItemNoDiscount, 4), "0%")
'Else
'    xTotal_itemDiscount_rate.Caption = ""
'End If

'If Val(xDiscount_offer.Caption) <> 0 And nTotalItem <> 0 Then
'    xDiscount_offer_rate.Caption = Format(Val(xDiscount_offer.Caption) / nTotalItem, "0%")
'Else
'    xDiscount_offer_rate.Caption = ""
'End If

bIg = True
'If Val(xTotal_item_offer.Caption) <> 0 Then
'    If Round(Val(xDiscount_add_rate.text), nRound) <> Round(Val(xDiscount_add.text) / Val(xTotal_item_offer.Caption) * 100, nRound) Then
'        xDiscount_add_rate.text = Myvalue(Round((Val(xDiscount_add.text) / Val(xTotal_item_offer.Caption)) * 100, nRound))
'    End If
'Else
'    xDiscount_add_rate.text = ""
'End If

'If xIsOffer.Value = 1 Then
'    xDiscount.text = Val(xDiscount_offer.Caption) + Val(xDiscount_add.text)
'ElseIf xIsOffer.Value = 0 And Val(xDiscount_add.text) <> 0 Then
'    'xDiscount.text = Myvalue(xDiscount_add.text)
'    xDiscount_add.text = 0
'End If

If Val(xtotal_item.Caption) <> 0 Then
    If Round(Val(xrate.Caption), nRound) <> Round(Val(xdiscount.Caption) / Val(xtotal_item.Caption) * 100, nRound) Then
        xrate.Caption = Myvalue(Round((Val(xdiscount.Caption) / Val(xtotal_item.Caption)) * 100, nRound))
    End If
Else
    xrate.Caption = ""
End If

xtotal.Caption = mRound(nTotalItem - Val(xdiscount.Caption), 2)
bIg = False
End With
End Function
Private Sub CardLookup(Optional pFilter As String = "")
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(5, 1)
Set Generalarray(0) = Me

Generalarray(1) = "SELECT TOP 1000 " & _
                  " FILE6_90BH.DOC_NO ," & _
                  " FORMAT(FILE6_90BH.date,'yyyy/M/d')," & _
                  " CONVERT(VARCHAR(10), FILE6_90BH.time, 108)," & _
                  " FROM FILE6_90BH"

If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & " WHERE " & cFilter
End If

Generalarray(2) = "Order by DATE DESC"
Generalarray(3) = 6000
Generalarray(5) = False

listarray(0, 0) = "—ﬁ„ «·„” ‰œ"
listarray(0, 1) = "(FILE6_90BH.DOC_NO = 'cFilter' OR FILE6_90BH.DOC_NO2 = 'cFilter')"

listarray(1, 0) = "«· «—ÌŒ"
listarray(1, 1) = "(##FILE6_90BH.DATE##)"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(1, 1) = 1200

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.Caption = "«” ⁄·«„"
oSearchDoc.Show 1
End Sub
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
For i = 1 To grid1.Rows - 1
    grid1.TextMatrix(i, 0) = i
Next
End Sub
Private Sub Fixgrd()
With grid1
'                 0          1          2         3        4          5              6          7          8            9           10          11             12           13            14            15          16          17           18           19        20
.FormatString = "„.|" & "»«—ﬂÊœ|" & "„Ê”„|" & "„’‰⁄|" & "„ﬂ »|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "„ﬁ«”|" & "«··Ê‰|" & "”⁄— „” Â·ﬂ|" & "⁄œœ|" & "”⁄— «·»Ì⁄|" & "‰”»… Œ’„|" & "«·≈Ã„«·Ï|" & "«·—’Ìœ|" & "ﬂÊœ «·»«∆⁄|" & "«·»«∆⁄|" & "«· ﬂ·›…|" & "„” ‰œ «·⁄—÷|" & "—ﬁ„ «·⁄—÷|" & "Œ’„ «·⁄—÷|"
.RowHeight(0) = 600
.WordWrap = True


'.ColHidden(14) = True

.ColWidth(0) = 600
.ColWidth(1) = 1000
.ColWidth(2) = 0
.ColWidth(3) = 1200
.ColWidth(4) = 0
.ColWidth(5) = 1300
.ColWidth(6) = 2500
.ColWidth(7) = 700
.ColWidth(8) = 1000
.ColWidth(9) = 1000
.ColWidth(10) = 800
.ColWidth(11) = 1000
.ColWidth(12) = 600
.ColWidth(13) = 1200
.ColWidth(16) = 1700
'.ColHidden(12) = True
.ColHidden(15) = True
.ColHidden(17) = True
.ColHidden(18) = True
.ColHidden(19) = True
.ColHidden(20) = True

.ColHidden(.Cols - 2) = True
.ColHidden(.Cols - 1) = True
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
.Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = flexAlignRightCenter

Dim con As New ADODB.Connection
sOk = openCon(con)
Dim bGroup As Boolean
Dim sGroup As String
For i = 1 To grid1.Rows - 2
    If .ValueMatrix(i, 9) <> .ValueMatrix(i, 11) Then
        .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = vbRed
        .Cell(flexcpFontUnderline, i, 1, i, .Cols - 1) = vbRed
    End If
    If sOk = "ok" And IsDgt(grid1.TextMatrix(i, 1)) Then
        grid1.TextMatrix(i, 14) = fnBalance(grid1.TextMatrix(i, 1), con, sStore, , grid1.TextMatrix(grid1.Row, grid1.Cols - 1))
    End If
    If .ValueMatrix(i, 19) > 0 Then
        If sGroup <> .TextMatrix(i, 19) Then
            sGroup = .TextMatrix(i, 19)
            bGroup = Not bGroup
        End If
        .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = RGB(IIf(bGroup, 255, 200), 255, 210)
    Else
        .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = vbWhite
    End If
Next
closeCon con
'.ShowCell .Rows - 1, 0
End With
End Sub
Private Function RemoveItem(Row As Long) As Boolean
If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel) <> vbOK Then
    Exit Function
End If

If grid1.TextMatrix(Row, grid1.Cols - 1) <> "" Then
    On Error GoTo myerror
    Dim con As New ADODB.Connection
    openCon con
    con.Execute "Delete  From FILE6_90B where id = " & grid1.TextMatrix(Row, grid1.Cols - 1)
    closeCon con
End If
grid1.RemoveItem Row
RemoveItem = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
closeCon con
End Function
Private Function myreplaceGrd(Row, con As ADODB.Connection) As Boolean
Dim aInsert As Variant
With grid1
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
        aInsert = AddFlag(Empty, "ITEM", addstring(grid1.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "QUANT", .ValueMatrix(i, 10))
        aInsert = AddFlag(aInsert, "PRICE", .ValueMatrix(i, 11))
        aInsert = AddFlag(aInsert, "PRICE_c", .ValueMatrix(i, 9))
        aInsert = AddFlag(aInsert, "COST", .ValueMatrix(i, 17))
        aInsert = AddFlag(aInsert, "MAN", addstring(.TextMatrix(i, 15)))
        aInsert = AddFlag(aInsert, "S_OKAZ", .ValueMatrix(i, 12))
        aInsert = AddFlag(aInsert, "USER_IP", addstring(cIpName))
        aInsert = AddFlag(aInsert, "DOC_OFFER", addstring(grid1.TextMatrix(i, 18)))
        aInsert = AddFlag(aInsert, "OFFER_NO", addstring(grid1.TextMatrix(i, 19)))
        
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            aInsert = AddFlag(aInsert, "DOC_NO", addstring(xdoc_no.Caption))
            con.Execute addInsert(aInsert, "FILE6_90B")
        Else
            con.Execute addUpdate(aInsert, "FILE6_90B", "ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
myreplaceGrd = True
End Function
Public Function myReplaceRefund(pDiscount_add As Double, pDiscount_total As Double) As Boolean
Dim con As New ADODB.Connection
openCon con

Dim aInsert As Variant
With oSalesRefund
Dim sdoc_no_new As String

sdoc_no_new = IncData("FILE6_90BH", "DOC_NO", con)
If sdoc_no_new = "" Then sdoc_no_new = "1"

aInsert = AddFlag(Empty, "DOC_NO", sdoc_no_new)
aInsert = AddFlag(aInsert, "[DISCOUNT]", -1 * Val(.xdiscount_ret.Caption))
aInsert = AddFlag(aInsert, "[ORDER_NO]", addstring(sOrder_No))
aInsert = AddFlag(aInsert, "[DATE]", addDate(Date))
aInsert = AddFlag(aInsert, "[TYPE]", "1")
aInsert = AddFlag(aInsert, "[USERNAME]", addstring(cUserName))
aInsert = AddFlag(aInsert, "[USER_IP]", addstring(GetComputerName))
aInsert = AddFlag(aInsert, "[SALES_RET]", addstring(.xdoc_no.Caption))
End With

con.BeginTrans
On Error GoTo myerror
con.Execute addInsert(aInsert, "FILE6_90BH")

With oSalesRefund.grid1
For i = 1 To .Rows - 2
    aInsert = AddFlag(Empty, "ITEM", addstring(.TextMatrix(i, 1)))
    aInsert = AddFlag(aInsert, "QUANT", -1 * .ValueMatrix(i, 8))
    aInsert = AddFlag(aInsert, "PRICE", .ValueMatrix(i, 9))
    aInsert = AddFlag(aInsert, "PRICE_C", .ValueMatrix(i, 7))
    aInsert = AddFlag(aInsert, "COST", .ValueMatrix(i, 15))
    aInsert = AddFlag(aInsert, "MAN", addstring(.TextMatrix(i, 15 + 1)))
    aInsert = AddFlag(aInsert, "USER_IP", addstring(GetComputerName))
    aInsert = AddFlag(aInsert, "OFFER_NO", addstring(.TextMatrix(i, 11)))
    aInsert = AddFlag(aInsert, "DOC_OFFER", addstring(.TextMatrix(i, 12)))
    aInsert = AddFlag(aInsert, "DISCOUNT_OFFER", -1 * .ValueMatrix(i, 13))
    aInsert = AddFlag(aInsert, "DOC_NO", addstring(sdoc_no_new))
    con.Execute addInsert(aInsert, "FILE6_90B")
Next
End With
con.CommitTrans
closeCon con

myReplaceRefund = True
bUpdated = True

MsgBox " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"


Unload oSalesRefund

If Not openCardTable(tbMode.tbFind, xdoc_no.Caption) Then
    If Not openCardTable Then
        myDefine
    End If
End If
Finally:
closeCon con
Exit Function
myerror:
MsgBox Err.Description
If TransCount(con) > 0 Then con.RollbackTrans
Err.Clear
GoTo Finally
End Function
Private Sub myLoadGrd()
With grid1
Dim cString As String
cString = "SELECT FILE6_90B.ITEM," & _
          "FILE1_10.MOSM," & _
          "FACT.DESCA," & _
          "FILE1_10.SUPP," & _
          "FILE1_10.MODELFACT0," & _
          "FILE1_10.DESCA," & _
          "FILE1_10.SCAL," & _
          "FILE1_10.COLOR," & _
          "FILE6_90B.PRICE_C," & _
          "FILE6_90B.QUANT," & _
          "FILE6_90B.PRICE," & _
          "FILE6_90B.S_OKAZ," & _
          "FILE6_90B.TOTAL," & _
          " 0 AS BALANCE," & _
          "FILE6_90B.MAN," & _
          "FILE6_25.DESCA," & _
          "FILE6_90B.COST," & _
          "FILE6_90B.DOC_OFFER," & _
          "FILE6_90B.OFFER_NO," & _
          "FILE6_90B.DISCOUNT_OFFER," & _
          "FILE6_90B.ID " & _
          " FROM FILE6_90B " & _
          " INNER JOIN FILE1_10 ON FILE6_90B.ITEM = FILE1_10.ITEM" & _
          " LEFT JOIN FACT ON FACT.CODE = FILE1_10.FACT" & _
          " LEFT JOIN FILE6_25 ON FILE6_90B.MAN  = FILE6_25.CODE "
cString = cString & " WHERE FILE6_90B.DOC_NO = " & MyParn(xdoc_no.Caption)
cString = cString & " ORDER BY FILE6_90B.ID"
Set grid1.DataSource = myRs(cString)
myAddItem
End With
CalcTotals
Fixgrd
End Sub
Private Function mysave(Optional bEnd As Boolean = True, Optional bPrint As Boolean = True, Optional bSendReceipt As Boolean = False) As Boolean

If Not myValid Then Exit Function
CalcTotals
If Not myreplace Then Exit Function
End Function
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = True) As Boolean
With grid1
If xtype.Tag = "" Then Exit Function
If Not IsNumeric(.TextMatrix(Row, 1)) Then Exit Function
If .ValueMatrix(Row, 10) = 0 Then Exit Function
If Not IsNumeric(.TextMatrix(Row, 11)) Then Exit Function
End With
validRow = True
End Function
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal col As Long)
KeyCode = 0
If bStopCell Then
    bStopCell = False
ElseIf col = 1 And grid1.TextMatrix(Row, col) = "" Then
ElseIf col < 10 Then
    grid1.col = 10
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 1, 10)
    grid1.ShowCell grid1.Row, 0
End If
End Sub
Private Sub myAddItem()
grid1.AddItem ""
MakeSerial
End Sub
Private Sub myUndo()
If xdoc_no.Tag = DefineMode Then
    If Not openCardTable Then
        CmdNewInv_Click
    End If
Else
    If Not openCardTable(tbMode.tbFind, xdoc_no.Caption) Then
        If Not openCardTable Then
            myDefine
        End If
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CmdNext_Click()
If Not openCardTable(tbMode.tbNext, xdoc_no.Caption) Then
    If Not openCardTable(tbMode.tblast) Then
        myDefine
    End If
End If
End Sub
Private Sub CmdPrevious_Click()
If Not openCardTable(tbMode.tbPrevious, xdoc_no.Caption) Then
    If Not openCardTable(tbMode.tbFirst) Then
        myDefine
    End If
End If
End Sub
Private Sub CmdFirst_Click()
If Not openCardTable(tbMode.tbFirst) Then
    myDefine
End If
End Sub
Private Sub CmdLast_Click()
If Not openCardTable Then
    myDefine
End If
End Sub
Private Function openCardTable(Optional pMode As Integer = tbMode.tblast, Optional pDoc_no As String = "", Optional pDoc_no_Filter As String = "") As Boolean
Dim cString As String
Dim cWhere As String
Dim cOrder As String
Dim nRecords As Long, nRecord As Long

If (pMode = tbMode.tbFind Or pMode = tbMode.tbNext Or pMode = tbMode.tbPrevious) And Trim(pDoc_no) = "" And Trim(pDoc_no_Filter) = "" Then
    Exit Function
End If

Me.MousePointer = vbHourglass

On Error GoTo myerror
cFilter = retFilter

Set CardTable = New ADODB.Recordset
cString = "SELECT TOP 1 FILE6_90BH.*," & _
          " FILE6_90BH.INV_NO AS DOC_NO_SALES," & _
          " ONLINE_TYPE_CODES.DESCA AS TYPE_DESCA" & _
          " FROM FILE6_90BH " & _
          " INNER JOIN ONLINE_TYPE_CODES ON FILE6_90BH.[TYPE] = ONLINE_TYPE_CODES.CODE"
  
If pMode = tbMode.tbFirst Then
    cOrder = "Order by FILE6_90BH.DOC_NO"
ElseIf pMode = tbMode.tblast Then
    cOrder = "Order by FILE6_90BH.DOC_NO DESC"
ElseIf pMode = tbMode.tbFind Then
    cWhere = "FILE6_90BH.Doc_no = " & MyParn(pDoc_no)
ElseIf pMode = tbMode.tbPrevious Then
    cWhere = "FILE6_90BH.DOC_NO < " & MyParn(pDoc_no)
    cOrder = "Order by FILE6_90BH.DOC_NO DESC"
ElseIf pMode = tbMode.tbNext Then
    cWhere = "FILE6_90BH.DOC_NO > " & MyParn(pDoc_no)
    cOrder = "Order by FILE6_90BH.DOC_NO"
End If

If cWhere <> "" Then
    cString = cString & " WHERE " & cWhere
End If

If cFilter <> "" Then
    cString = cString & IIf(cWhere = "", " WHERE ", " AND ") & cFilter
End If

cString = cString & " " & cOrder

Set CardTable = myRs(cString)

If (Not CardTable.EOF) Then
    myload
    openCardTable = True
End If
Me.MousePointer = vbNormal
Exit Function
myerror:
Me.MousePointer = vbNormal
MsgBox Err.Description
Err.Clear
End Function
Private Function validQuant() As Boolean
Dim nType As Integer
For i = 1 To grid1.Rows - 1
    If (xtype.Tag = 1 Or xtype.Tag = 3) And grid1.ValueMatrix(i, 10) < 0 Then
        Exit Function
    ElseIf xtype.Tag = 2 Or xtype.Tag = 4 And grid1.ValueMatrix(i, 10) > 0 Then
        Exit Function
    ElseIf grid1.ValueMatrix(i, 10) = 0 Then
        Exit Function
    End If
    
    
'    If grid1.ValueMatrix(i, 10) = 0 Then
'    ElseIf nType = 0 Then
'        nType = IIf(grid1.ValueMatrix(i, 10) > 0, 1, 2)
'    Else
'        If nType <> IIf(grid1.ValueMatrix(i, 10) > 0, 1, 2) Then Exit Function
'    End If
Next
validQuant = True
End Function
Private Function retRecords(pDoc_no, ByRef nRecords As Long, ByRef nRecord As Long) As Variant
Dim cString As String, loctable As New ADODB.Recordset
If pDoc_no <> "" Then
    cString = "SELECT Count(FILE6_90BH.DOC_NO) AS records,COUNT(CASE WHEN FILE6_90BH.DOC_NO <= " & MyParn(pDoc_no) & " THEN 1 END) AS record"
Else
    cString = "SELECT Count(FILE6_90BH.DOC_NO) AS records,0 as record"
End If

cString = cString & " FROM FILE6_90BH"

If cFilter <> "" Then
    cString = cString & " WHERE " & cFilter
End If

Set loctable = myRs(cString)
If Not loctable.EOF Then
    nRecords = loctable!RECORDS
    nRecord = Val(loctable!Record & "")
End If
End Function
Private Function retFilter()
retFilter = "FILE6_90BH.ORDER_NO = " & MyParn(sOrder_No)
End Function
Private Sub xdoc_ret_GotFocus()
myGotFocus xdoc_ret
End Sub
Private Sub xdoc_ret_LostFocus()
myLostFocus xdoc_ret
End Sub

Private Sub xNotes_GotFocus()
myGotFocus xNotes
End Sub
Private Sub xNotes_LostFocus()
myLostFocus xNotes
End Sub
Private Sub xdate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xDate_LostFocus()
myLostFocus xDate
myValidDate xDate
End Sub
Private Sub HandleInit()
If cBranch = "00" Then
    nUser = enUser.MainUser
Else
    nUser = enUser.branch
End If
cmdSend.Visible = nUser = enUser.branch
End Sub
Private Function myReplaceOnline() As String
Dim sDoc_New As String
Dim sDate As String
Dim cAddress As String
Dim sMan As String

Dim cString As String

cString = "SELECT DSALES FROM DSALES WHERE BRANCH = " & MyParn(cBranch)
sDate = myFormat(rsValue(cString))
                                              
cString = " SELECT FILE6_90BH.DOC_NO," & _
          " FILE6_90BH.ORDER_NO," & _
          " FILE6_90BH.SALES_RET," & _
          " FILE6_90BH.TYPE," & _
          " FILE6_90H.NAME," & _
          " FILE6_90H.PHONE," & _
          " FILE6_90H.PAYMENT_ID," & _
          " FILE6_90H.Shipping_City," & _
          " FILE6_90H.CITY," & _
          " FILE6_90H.STREET, " & _
          " FILE6_90H.MAN, " & _
          " FILE6_90BH.DISCOUNT " & _
          " FROM FILE6_90BH " & _
          " INNER JOIN FILE6_90H ON FILE6_90BH.ORDER_NO = FILE6_90H.DOC_NO" & _
          " WHERE FILE6_90BH.DOC_NO = " & MyParn(xdoc_no.Caption)
       
Dim loctable As New ADODB.Recordset
Set loctable = myRs(cString)

If loctable.EOF And loctable.BOF Then
    MsgBox "«·„” ‰œ €Ì— „”Ã·"
    Exit Function
End If

If IsNull(loctable!MAN) Then
    sMan = "0001"
Else
    sMan = loctable!MAN
End If

aInsert = AddFlag(Empty, "CODE", addstring("0000"))
aInsert = AddFlag(aInsert, "ONLINE", "2")
aInsert = AddFlag(aInsert, "[Date]", addDate(sDate))
aInsert = AddFlag(aInsert, "STORE", addstring(cBranchStore))
aInsert = AddFlag(aInsert, "BOX", addstring(cManBox))

aInsert = AddFlag(aInsert, "DISCOUNT", Val(loctable!discount))
If Val(loctable!discount & "") <> 0 Then
    aInsert = AddFlag(aInsert, "username_disc", addstring("«Ê‰·«Ì‰"))
End If

aInsert = AddFlag(aInsert, "BRANCH", addstring(cBranch))
aInsert = AddFlag(aInsert, "PHONE", addstring(loctable!phone))
aInsert = AddFlag(aInsert, "MAN", addstring(loctable!MAN))

aInsert = AddFlag(aInsert, "INV_TYPE_ONLINE", loctable!Type)
aInsert = AddFlag(aInsert, "DOC_NO_ONLINE", loctable!DOC_NO)
aInsert = AddFlag(aInsert, "ONLINE_DOC", addstring(loctable!ORDER_NO))
aInsert = AddFlag(aInsert, "SALES_RET", addstring(loctable!SALES_RET))
    
aInsert = AddFlag(aInsert, "PAYMENT_ID", addstring(loctable!PAYMENT_ID))
aInsert = AddFlag(aInsert, "userName", addstring(cUserName))
aInsert = AddFlag(aInsert, "USER_IP", addstring(GetComputerName))

cAddress = loctable!Shipping_City & ""
If loctable!Street & "" <> "" Then
    cAddress = cAddress & Tr(cAddress, " ") & loctable!Street
End If
aInsert = AddFlag(aInsert, "ADDRESS", addstring(cAddress))
aInsert = AddFlag(aInsert, "[NAME]", addstring(loctable!Name))
aInsert = AddFlag(aInsert, "[city]", addstring(loctable!city))
aInsert = AddFlag(aInsert, "[Shipping_City]", addstring(loctable!Shipping_City))
aInsert = AddFlag(aInsert, "[street]", addstring(loctable!Street))
'aInsert = AddFlag(aInsert, "[Payment_Method]", addstring(loctable!Payment_Method))



sDoc_New = NewflagDocRs(sDate, cManBox)
aInsert = AddFlag(aInsert, "[doc_no]", addstring(sDoc_New))

Dim con As New ADODB.Connection
If openCon(con) <> "ok" Then
    GoTo Finally
End If


con.BeginTrans
con.Execute addInsert(aInsert, "FILE6_20H")

Dim cInsert As New ChilkatStringBuilder
cInsert.Append "insert into FILE6_20(" & _
                "DOC_NO," & _
                "ITEM," & _
                "QUANT," & _
                "PRICE," & _
                "PRICE_C," & _
                "[MAN]," & _
                "[USER_IP]," & _
                "[COST]," & _
                "[DOC_OFFER]," & _
                "[OFFER_NO]," & _
                "[DISCOUNT_OFFER]" & _
                ")"
With grid1
If grid1.Rows > 2 Then
    cInsert.Append " VALUES "
    For i = 1 To .Rows - 2
        cInsert.Append "("
        cInsert.Append addstring(sDoc_New) & ","
        cInsert.Append addvalue(.TextMatrix(i, 1)) & ","
        cInsert.Append .ValueMatrix(i, 10) & ","
        cInsert.Append .ValueMatrix(i, 11) & ","
        cInsert.Append .ValueMatrix(i, 9) & ","
        cInsert.Append addstring(sMan) & ","
        cInsert.Append addstring(GetComputerNamecIpName) & ","
        cInsert.Append addstring(.ValueMatrix(i, 17)) & ","
        cInsert.Append addstring(.TextMatrix(i, 18)) & ","
        cInsert.Append addstring(.TextMatrix(i, 19)) & ","
        cInsert.Append addstring(.TextMatrix(i, 20)) & ")"
        cInsert.Append ","
    Next
    cInsert.Shorten 1
End If
End With

con.Execute cInsert.GetAsString()
con.Execute "UPDATE FILE6_90BH " & _
            " SET FILE6_90HB.INV_NO = " & MyParn(sDoc_New) & _
            " WHERE FILE6_90BH = " & MyParn(xdoc_no.Caption)
UpdateInvTotal sDoc_New, con

con.CommitTrans

myReplaceOnline = sDoc_New
Finally:
    closeCon con
Exit Function
myerror:
MsgBox Err.Description
If TransCount(con) > 0 Then con.RollbackTrans
Err.Clear
sdoc_no_new = ""
GoTo Finally
End Function
Private Sub xdate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xdate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xSales_replace_Click()
ItemsLook Me, oSearchItem, , , "FILE1_10.ITEM IN(SELECT ITEM FROM FILE6_20 WHERE DOC_NO = " & MyParn(xSales_replace.Caption) & ")", False
End Sub
