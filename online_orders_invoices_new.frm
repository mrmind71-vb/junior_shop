VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form orders_online_invoices 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«’‰«› ÿ·»Ì« "
   ClientHeight    =   10575
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
   ScaleHeight     =   10575
   ScaleWidth      =   20400
   Visible         =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
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
      Left            =   225
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   9000
      Width           =   3570
      Begin Threed.SSCommand cmdFirst 
         Height          =   510
         Left            =   2700
         TabIndex        =   19
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
         Picture         =   "online_orders_invoices_new.frx":0000
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "online_orders_invoices_new.frx":21A7
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   510
         Left            =   1800
         TabIndex        =   20
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
         Picture         =   "online_orders_invoices_new.frx":41EE
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "online_orders_invoices_new.frx":62D9
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   510
         Left            =   900
         TabIndex        =   21
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
         Picture         =   "online_orders_invoices_new.frx":82D3
         Caption         =   "·«Õﬁ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "online_orders_invoices_new.frx":A3E4
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   510
         Left            =   0
         TabIndex        =   22
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
         Picture         =   "online_orders_invoices_new.frx":C3DE
         Caption         =   "√ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "online_orders_invoices_new.frx":E602
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
      Height          =   1770
      Left            =   7425
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   0
      Width           =   12885
      Begin VB.TextBox xValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   945
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.TextBox xDate_mail 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   585
         Width           =   2085
      End
      Begin VB.TextBox xPayment_id 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   8055
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1305
         Width           =   3615
      End
      Begin VB.TextBox xCharge2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3645
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   945
         Width           =   2040
      End
      Begin VB.CheckBox xClosed 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„€·ﬁ…"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7470
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   585
         Width           =   780
      End
      Begin VB.TextBox xNotes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1305
         Width           =   5550
      End
      Begin VB.TextBox xShip_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3645
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   2040
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   9585
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   945
         Width           =   2085
      End
      Begin MSDataListLib.DataCombo xship 
         Height          =   360
         Left            =   3645
         TabIndex        =   3
         Top             =   540
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand xtype 
         Height          =   330
         Left            =   9585
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   585
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
         Caption         =   "«Œ «— ‰Ê⁄ «·ÿ·»Ì…"
         TagVariant      =   "«Œ «— ‰Ê⁄ «·ÿ·»Ì…"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdSavePayment 
         Height          =   330
         Left            =   7425
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1305
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   582
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "save"
         TagVariant      =   "«Œ «— ‰Ê⁄ «·ÿ·»Ì…"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin VB.Label lblShip 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ﬁÌ„… «·»Ê·Ì’…"
         Height          =   285
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   990
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   " «—ÌŒ „Ì·"
         Height          =   330
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   540
         Width           =   915
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ ”œ«œ"
         Height          =   330
         Left            =   11835
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   1350
         Width           =   825
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„’«—Ì› «·‘Õ‰"
         Height          =   285
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   945
         Width           =   1230
      End
      Begin VB.Label xStage 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   7425
         TabIndex        =   60
         Tag             =   "1"
         Top             =   945
         Width           =   2130
      End
      Begin VB.Label xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   135
         TabIndex        =   49
         Tag             =   "1"
         Top             =   225
         Width           =   2085
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   " «—ÌŒ  ”·Ì„/»Ì⁄"
         Height          =   330
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "‘—ﬂ… «·‘Õ‰"
         Height          =   270
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   585
         Width           =   945
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "‰Ê⁄ «·ÿ·»Ì…"
         Height          =   330
         Left            =   11745
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   585
         Width           =   1050
      End
      Begin VB.Label xdoc_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   9585
         TabIndex        =   45
         Tag             =   "1"
         Top             =   225
         Width           =   2085
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ „” ‰œ"
         Height          =   330
         Left            =   11745
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   225
         Width           =   825
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„·ÕÊŸ…"
         Height          =   285
         Left            =   5850
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   1350
         Width           =   690
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "»Ê·Ì’… «·‘Õ‰"
         Height          =   285
         Left            =   5805
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«· «—ÌŒ"
         Height          =   330
         Left            =   11790
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   945
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
      Left            =   2250
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   0
      Width           =   5145
      Begin Threed.SSCommand cmdNewInv 
         Height          =   510
         Left            =   2655
         TabIndex        =   24
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
         Picture         =   "online_orders_invoices_new.frx":106D3
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_invoices_new.frx":126DB
      End
      Begin Threed.SSCommand cmddel 
         Height          =   510
         Left            =   1350
         TabIndex        =   25
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
         Picture         =   "online_orders_invoices_new.frx":14692
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_invoices_new.frx":16E2E
      End
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   3915
         TabIndex        =   23
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
         Picture         =   "online_orders_invoices_new.frx":192C2
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_invoices_new.frx":1B68D
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   26
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
         Picture         =   "online_orders_invoices_new.frx":1D736
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
      Left            =   5850
      RightToLeft     =   -1  'True
      TabIndex        =   31
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
         Picture         =   "online_orders_invoices_new.frx":1FA59
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_invoices_new.frx":2237E
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
         Picture         =   "online_orders_invoices_new.frx":24BD2
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_invoices_new.frx":26D32
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
      TabIndex        =   27
      Top             =   10200
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
         TabIndex        =   28
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
         TabIndex        =   29
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
         TabIndex        =   30
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
      Height          =   4335
      Left            =   225
      TabIndex        =   13
      Top             =   1800
      Width           =   20085
      _cx             =   35428
      _cy             =   7646
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
   Begin VSFlex7Ctl.VSFlexGrid grdError 
      Height          =   2805
      Left            =   225
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6165
      Width           =   20085
      _cx             =   35428
      _cy             =   4948
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
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
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
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   14130
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   8955
      Width           =   6180
      Begin VB.TextBox xDiscount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00400000&
         Height          =   330
         Left            =   3870
         MaxLength       =   15
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   585
         Width           =   1005
      End
      Begin VB.TextBox xRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00400000&
         Height          =   330
         Left            =   3150
         MaxLength       =   10
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   585
         Width           =   690
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "≈Ã„«·Ì «·«’‰«›"
         Height          =   285
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   225
         Width           =   1320
      End
      Begin VB.Label label50 
         BackColor       =   &H00FFFFFF&
         Caption         =   "’«›Ì «·ﬁÌ„…"
         Height          =   285
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   630
         Width           =   1185
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·Œ’„"
         Height          =   285
         Left            =   4950
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   630
         Width           =   735
      End
      Begin VB.Label xtotal_item 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   90
         TabIndex        =   54
         Top             =   225
         Width           =   1365
      End
      Begin VB.Label xtotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   3870
         TabIndex        =   53
         Top             =   225
         Width           =   1005
      End
      Begin VB.Label xtotalQuant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   90
         TabIndex        =   52
         Top             =   585
         Width           =   1365
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "≈Ã„«·Ì «·ﬂ„Ì…"
         Height          =   285
         Left            =   4950
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   3825
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Top             =   8955
      Width           =   4290
      Begin Threed.SSCommand cmdEdit 
         Height          =   555
         Left            =   45
         TabIndex        =   18
         Top             =   180
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   979
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
         Caption         =   " „ «· ⁄œÌ·"
         ButtonStyle     =   2
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_invoices_new.frx":2901F
      End
      Begin Threed.SSCommand cmdMsg 
         Height          =   555
         Left            =   2115
         TabIndex        =   17
         Top             =   180
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   979
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
         Caption         =   "«—”«· —”«·…"
         ButtonStyle     =   2
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_orders_invoices_new.frx":2B315
      End
   End
   Begin VB.Frame fmSend 
      BackColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   2205
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   720
      Width           =   3615
      Begin VB.Label lblReturn 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "›« Ê—… «·„— Ã⁄"
         Height          =   270
         Left            =   2115
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   630
         Width           =   1110
      End
      Begin VB.Label xSales_ret 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   90
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
         Top             =   270
         Width           =   1215
      End
   End
   Begin VB.Frame fmReplace 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   2250
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   720
      Width           =   3570
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„‰œÊ»"
         Height          =   270
         Left            =   2025
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   630
         Width           =   600
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "›« Ê—… «” »œ«·"
         Height          =   270
         Left            =   2025
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   270
         Width           =   1080
      End
      Begin VB.Label xMan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   45
         TabIndex        =   61
         Tag             =   "1"
         Top             =   585
         Width           =   1905
      End
      Begin VB.Label xSales_Replace 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   45
         TabIndex        =   59
         Tag             =   "1"
         Top             =   225
         Width           =   1905
      End
   End
End
Attribute VB_Name = "orders_online_invoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sMsg As String
Public sOrder_no As String
Public sSales_doc As String
Public sStore As String
Public myForm As Form
Public bEdit As Boolean
Dim dbm As New DBManager

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
Dim oSearchType As New Search_rs
Public sStage As String
Public bOpenOnly As Boolean
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
aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(xDiscount.text))
aInsert = AddFlag(aInsert, "[TYPE]", addvalue(xtype.Tag))
aInsert = AddFlag(aInsert, "[SHIP]", addstring(xship.BoundText))
aInsert = AddFlag(aInsert, "[SHIP_NO]", addstring(xShip_no.text))
aInsert = AddFlag(aInsert, "[ORDER_NO]", addstring(sOrder_no))
aInsert = AddFlag(aInsert, "[SALES_REPLACE]", addstring(xSales_Replace.Caption))
aInsert = AddFlag(aInsert, "[PAYMENT_ID]", addstring(xPayment_id.text))
aInsert = AddFlag(aInsert, "DATE_MAIL", addDate(xDate_mail.text))
aInsert = AddFlag(aInsert, "[DATE]", addDate(xDate.text))
aInsert = AddFlag(aInsert, "[charge2]", Val(xCharge2.text))
aInsert = AddFlag(aInsert, "[VALUE]", Val(xValue.text))
aInsert = AddFlag(aInsert, "[NOTES]", addstring(xNotes.text))
'aInsert = AddFlag(aInsert, "[DATE1]", addDate(xdate1.Caption))

On Error GoTo myerror
Dim con As New ADODB.Connection
If OpenCon(con) <> "ok" Then Exit Function
con.BeginTrans
If xdoc_no.Tag = DefineMode Then
    xdoc_no.Caption = Newflag("FILE6_90BH", "DOC_NO", con)
    aInsert = AddFlag(aInsert, "DOC_NO", addstring(xdoc_no.Caption))
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(cUserName))
    aInsert = AddFlag(aInsert, "[USER_IP]", addstring(GetComputerName))
    aInsert = AddFlag(aInsert, "[STAGE]", IIf(xtype.Tag = 1 Or xtype.Tag = 11, "1", "7"))
    con.Execute addInsert(aInsert, "FILE6_90BH")
Else
    con.Execute addUpdate(aInsert, "FILE6_90BH", "DOC_NO = " & addstring(xdoc_no.Caption))
End If

If bOffer Or xtype.Tag = "13" Then
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
        CellPos 13, grid1.Row, grid1.Col
    End If
    GrdDesc grid1.TextMatrix(grid1.Row, 1), grid1.Row
    Grid1_AfterEdit grid1.Row, grid1.Col
ElseIf ActiveControl.Name = xSales_Replace.Name Then
    grid1.Row = grid1.Rows - 1
    grid1.TextMatrix(grid1.Row, 1) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    If Not bNew Then
        CellPos 13, grid1.Row, grid1.Col
    End If
    GrdDesc grid1.TextMatrix(grid1.Row, 1), grid1.Row
    Grid1_AfterEdit grid1.Row, grid1.Col
ElseIf ActiveControl.Name = cmdInform.Name Then
    xdoc_no.Caption = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    Unload oSearchDoc
    myUndo
ElseIf sControl = "inv_replace" Then
    If Not myValid() Then Exit Sub
    sDoc = oSearchDocRet.grid1.TextMatrix(oSearchDocRet.grid1.Row, 0)
    Unload oSearchDocRet
    If myValid Then
        addRefundOrder sDoc, "11"
    End If
ElseIf sControl = "inv_ret" Then
    If Not myValid Then Exit Sub
    sDoc = oSearchDocRet.grid1.TextMatrix(oSearchDocRet.grid1.Row, 0)
    Unload oSearchDocRet
    If myValid Then
        addRefundOrder sDoc, "2"
    End If
ElseIf sControl = "inv_ret_all" Then
    If Not myValid(, True) Then Exit Sub
    sDoc = oSearchDocRet.grid1.TextMatrix(oSearchDocRet.grid1.Row, 0)
    Unload oSearchDocRet
    myReplaceRefundAll sDoc
ElseIf ActiveControl.Name = xtype.Name Then
    xSales_Replace.Caption = ""
    xSales_ret.Caption = ""
    xdoc_no_sales.Caption = ""
    xMan.Tag = ""
    xMan.Caption = ""
    
    fmReplace.Visible = False
    fmSend.Visible = False
                
    If Not myValid(, oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 0) = "3") Then Exit Sub
    
    If oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 0) = "2" Then
        Unload oSearchType
        addRefund "2"
    ElseIf oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 0) = "3" Then
        Unload oSearchType
        addRefundAll
    ElseIf oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 0) = "11" Then
        Unload oSearchType
        addRefund "11"
    ElseIf oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 0) = "12" Then
        addReplace
    ElseIf oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 0) = "1" Then
        addInv
    ElseIf oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 0) = "13" Then
        addCash
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Function addRefund(pType As String) As Integer
Dim aValues As Variant

aValues = rsValues("SELECT COUNT(*) AS COUNT_OF,MIN(DOC_NO) AS DOC_NO FROM FILE6_20H WHERE IS_RETURN = 0 AND ISCLOSED = 1 AND ONLINE_DOC = " & MyParn(sOrder_no))
If IsEmpty(aValues) Then
    MsgBox "·«  ÊÃœ ›Ê« Ì— "
    Exit Function
End If

If retFlag(aValues, "COUNT_OF") = 1 Then
    addRefundOrder retFlag(aValues, "DOC_NO") & "", pType
Else
    invLook "inv_ret"
End If
End Function
Private Function addRefundAll() As Integer
Dim aValues As Variant
aValues = rsValues("SELECT COUNT(*) AS COUNT_OF,MIN(DOC_NO) AS DOC_NO FROM FILE6_20H WHERE IS_RETURN = 0 AND ISCLOSED = 1 AND ONLINE_DOC = " & MyParn(sOrder_no))

If IsEmpty(aValues) Then
    MsgBox "·«  ÊÃœ ›Ê« Ì— "
    Exit Function
End If

If retFlag(aValues, "COUNT_OF") = 1 Then
    myReplaceRefundAll retFlag(aValues, "DOC_NO") & ""
Else
    invLook "inv_ret_all"
End If
End Function
Private Function addReplace() As Integer
xtype.Tag = oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 0)
xtype.Caption = oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 1)
Unload oSearchType

Dim aValues As Variant
aValues = rsValues("SELECT COUNT(*) AS COUNT_OF,MIN(DOC_NO) AS DOC_NO FROM FILE6_20H WHERE IS_RETURN = 0 AND  ONLINE_DOC = " & MyParn(sOrder_no))
If IsEmpty(aValues) Then
    MsgBox "·«  ÊÃœ ›Ê« Ì— ··«” »œ«·"
    Exit Function
End If

If retFlag(aValues, "COUNT_OF") = 1 Then
    xSales_Replace.Caption = retFlag(aValues, "DOC_NO") & ""
    Handlecontrols xdoc_no.Tag
Else
    invLook "inv_replace"
End If

fmReplace.Visible = True
End Function
Private Function addInv() As Integer
xtype.Tag = oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 0)
xtype.Caption = oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 1)
Unload oSearchType

Dim strSql As String
strSql = "SELECT TOP 1 SALES_RET " & _
         " FROM FILE6_90BH WHERE ORDER_NO = " & MyParn(sOrder_no) & _
         " AND TYPE = 2" & _
         " ORDER BY DATE DESC,DOC_NO DESC"

Dim sValue As String
sValue = rsValue(strSql) & ""
If sValue = "" Then
    MsgBox "·«  ÊÃœ ›Ê« Ì— ··«” »œ«·"
    Exit Function
End If
'
'If retFlag(aValues, "COUNT_OF") = 1 Then
    'xSales_Replace.Caption = retFlag(aValues, "DOC_NO") & ""
xSales_Replace.Caption = sValue
Handlecontrols xdoc_no.Tag
'Else
'    invLook "inv_add"
'End If

fmReplace.Visible = True
End Function
Private Function addCash() As Integer
xtype.Tag = oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 0)
xtype.Caption = oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 1)
Unload oSearchType

xSales_Replace.Caption = sValue
Handlecontrols xdoc_no.Tag
grid1.Enabled = False
grdError.Editable = False
lblShip.Visible = True
xValue.Visible = True
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
                  " WHERE FILE6_20H.ONLINE_DOC = " & MyParn(sOrder_no) & _
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
Private Function addRefundOrder(pDoc_no, pType As String) As Boolean
Dim locTable As New ADODB.Recordset
Dim cString As String
cString = "select TOP 1 FILE6_20H.BRANCH," & _
          " FILE6_20H.PRINTED," & _
          " FILE6_20H.ISINVOICE," & _
          " FILE6_20H.TOTAL_ITEM," & _
          " FILE6_20H.DATE, " & _
          " FILE6_20H.MAN, " & _
          " FILE6_20H.ISRET" & _
          " FROM FILE6_20H " & _
          " WHERE FILE6_20H.DOC_NO = " & MyParn(pDoc_no)

Set locTable = myRs(cString)
If locTable.EOF Then
    MsgBox "›« Ê—… €Ì— „”Ã·…"
    Exit Function
End If

If locTable!branch & "" = "00" Then
    MsgBox "›« Ê—… ›—⁄ —∆Ì”Ì"
    Exit Function
End If

If Not locTable!printed Then
    MsgBox "›« Ê—… „› ÊÕ…"
    Exit Function
End If

If locTable!ISINVOICE Then
    MsgBox "›« Ê—… „Ã„⁄…"
    Exit Function
End If

If locTable!TOTAL_ITEM <= 0 Then
    If locTable!TOTAL_ITEM = 0 Then
        MsgBox "›« Ê—… »œÊ‰ ﬁÌ„…"
    End If
    If locTable!TOTAL_ITEM < 0 Then
        MsgBox "›« Ê—… „— Ã⁄"
    End If
    Exit Function
End If

xMan.Tag = locTable!MAN & ""
Set oSalesRefund.myForm = Me
oSalesRefund.sDoc_no = pDoc_no
oSalesRefund.sFlag = pType
oSalesRefund.bOnline = True
oSalesRefund.Show 1
End Function

Private Sub CmdDel_Click()
If MsgBox("Õ–›", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub

Dim db As New clsDb

Dim inv_no As Variant
If Not db.rsValue("Select INV_NO from file6_90BH WHERE DOC_NO = " & xdoc_no.Caption, inv_no) Then
    GoTo CleanUp
End If

If Not IsNull(inv_no) Then
    MsgBox "›« Ê—… „»Ì⁄«  „”Ã·… ··„” ‰œ" & inv_no
    Exit Sub
End If

db.Sql = "Delete  From FILE6_90B where Doc_No = " & xdoc_no.Caption
db.Sql = "Delete  From FILE6_90BH where Doc_No = " & xdoc_no.Caption

If Not db.ExecuteTransaction Then GoTo CleanUp

Inform " „ «·Õ–› »‰Ã«Õ"

If Not openCardTable(tbMode.tbPrevious, xdoc_no.Caption) Then
    If Not openCardTable(tbMode.tbFirst) Then
        myDefine
    End If
End If
bUpdated = True
CleanUp:
Set db = Nothing
End Sub
Private Sub cmdEdit_Click()
If MsgBox(" „ «· ⁄œÌ·", vbDefaultButton2 + vbOKCancel) <> vbOK Then Exit Sub
    If sendEdit Then
        Inform_OK " „ «· ⁄œÌ·"
    
        If Not openCardTable(tbMode.tbFind, xdoc_no.Caption) Then
            If Not openCardTable Then
                myDefine
            End If
        End If
    End If
End Sub
Private Function sendEdit() As Boolean
    On Error GoTo myerror
    Dim con As New ADODB.Connection
    If opencn(con) Then
    
        sMsg = " „ «· ⁄œÌ·"
        Set online_error.myForm = Me
        online_error.sCaption = "»Ì«‰ «· ⁄œÌ·"
        online_error.Show 1
        
        If sMsg = "" Then
            Inform_OK "«·»Ì«‰ ÷—Ê—Ì"
            Exit Function
        End If
        
        Dim aInsert As Variant
        aInsert = AddFlag(Empty, "DESCA", addstring(sMsg))
        aInsert = AddFlag(aInsert, "USERNAME", addstring(cUserName))
        aInsert = AddFlag(aInsert, "ORDER_NO", addstring(sOrder_no))
        aInsert = AddFlag(aInsert, "DOC_NO_SUP", addvalue(xdoc_no.Caption))
        aInsert = AddFlag(aInsert, "[TIME]", "getdate()")
        
        con.BeginTrans
        con.Execute "UPDATE FILE6_90BH " & _
                    " SET STAGE = 8," & _
                    " TIME8 = getDate()" & _
                    " FROM FILE6_90BH WHERE DOC_NO = " & addstring(xdoc_no.Caption)
                    
        con.Execute addInsert(aInsert, "FILE6_90_MSG"), nAffect
        con.CommitTrans
    
    End If
    
    sendEdit = True
Finally:
    closeCon con
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
Resume Finally
End Function
Private Sub CmdExit_Click()
    sDoc_no = ""
    Unload Me
End Sub
Private Sub CmdInform_Click()
CardLookup
End Sub

Private Sub cmdMsg_Click()
SendMsg
End Sub
Private Sub SendMsg()
sMsg = ""
Set online_error.myForm = Me
online_error.sCaption = "»Ì«‰ «· ⁄œÌ·"
online_error.Show 1

If Trim(sMsg) = "" Then
    Inform "·„ Ì „  ”ÃÌ· «·—”«·…"
    Exit Sub
End If

Dim aInsert As Variant
aInsert = AddFlag(Empty, "DESCA", addstring(sMsg))
aInsert = AddFlag(aInsert, "USERNAME", addstring(cUserName))
aInsert = AddFlag(aInsert, "ORDER_NO", addstring(sOrder_no))
aInsert = AddFlag(aInsert, "DOC_NO_SUP", addvalue(xdoc_no.Caption))
aInsert = AddFlag(aInsert, "STAGE", addvalue(xStage.Tag))
aInsert = AddFlag(aInsert, "[TIME]", "getdate()")
On Error GoTo myerror
nAffrect = rsEx(addInsert(aInsert, "FILE6_90_MSG"))

myloadGrdError

Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub CmdNewInv_Click()
myDefine
'On Error Resume Next
'grid1.Select 1, 1
'Err.Clear
End Sub
Private Sub cmdSave_Click()
If grid1.Rows = 2 And xtype.Tag <> "13" Then Exit Sub
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
Private Sub cmdSaveDateMail_Click()
End Sub
Private Sub cmdSavePayment_Click()
If dbm.OpenCon Then
    If dbm.Execute("update file6_90BH " & _
                   " SET PAYMENT_ID = " & addstring(xPayment_id.text) & "," & _
                   " Ship_no = " & addstring(GetNumbersFromString(xShip_no.text)) & "," & _
                   " DATE_MAIL = " & addDate(xDate_mail.text) & _
                   " WHERE DOC_NO = " & MyParn(xdoc_no.Caption)) Then
        Inform " „ Õ›Ÿ —ﬁ„ «·”œ«œ »‰Ã«Õ"
    End If
    dbm.closeCon
End If
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
        'cmdType.SetFocus
        'OnlineTypeLookup Me, oSearchType, , , , "type"
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

Set xship.RowSource = myRs("SELECT CODE,DESCA FROM SHIP ORDER BY  STOPED,DESCA")
xship.ListField = "DESCA"
xship.BoundColumn = "CODE"

cDefClient = "0000"

If Not openCardTable Then
    myDefine
End If

Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Form_Resize()
'Dim nHeight As Long
'nHeight = Me.Height - (grid1.Top + 4100)
'If nHeight > 3000 Then
'    grid1.Height = nHeight
'End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
Set dbm = Nothing
Set orders_online_invoices = Nothing
If bUpdated And Not myForm Is Nothing Then
    myForm.myLoadCount
End If
Err.Clear
End Sub

Private Sub grdError_DblClick()
If grdError.Row > 0 Then
    online_Show_Msg.sMsg = grdError.TextMatrix(grdError.Row, 2)
    online_Show_Msg.Show 1
End If
End Sub

Public Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo myerror
Dim bNew As Boolean
With grid1
If Not myValid(True) Then
    On Error Resume Next
    grid1.SetFocus
    Err.Clear
    myLoadGrd
    If Row < grid1.Rows - 1 Then
        grid1.Select Row, Col
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
    If cBranch <> "00" Then
        .Editable = flexEDNone
    ElseIf Not bEditRecord Then
        .Editable = flexEDNone
    ElseIf xtype.Tag = "" Then
        .Editable = flexEDNone
    ElseIf .ValueMatrix(grid1.Row, 19) > 0 Then
        .Editable = flexEDNone
    ElseIf (.Col = 1 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "") Or (.Col = 10) Then
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
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    'If Col = 12 And grid1.TextMatrix(Row, Col) = "" Then Exit Sub
    CellPos KeyCode, Row, Col
End If
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim sBarCode As String
Dim nSal As Double
Dim nRet As Double
Dim cPassword As String, aUser As Variant
With grid1
    If Col = 10 Then
        If Not IsNumeric(.EditText) Then
            Cancel = True
            Exit Sub
        ElseIf Val(.EditText) = 0 Then
            MsgBox "€»— „”„ÊÕ »ﬂ„Ì… ’›—Ì…"
            Cancel = True
        ElseIf Val(.EditText) < 0 Then
            MsgBox "€»— „”„ÊÕ »ﬂ„Ì… €Ì— „ÊÃ»…"
        ElseIf xtype.Caption = "1" Or xtype.Caption = "3" Then
            nBalance = rsBalance(grid1.TextMatrix(grid1.Row, 1), cBranchStore, xDate.text, grid1.TextMatrix(grid1.Row, grid1.Cols - 1))
            If nBalance < Val(.EditText) Then
                Inform "«·—’Ìœ ·« Ì”„Õ"
                Exit Sub
            End If
        End If
    ElseIf Col = 1 Then
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
Private Function myValid(Optional bIgMsg As Boolean = False, Optional bIgShip As Boolean = False) As Boolean
If Not IsDate(xDate.text) Then
    If Not bIgMsg Then MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

'If Not bIgType Then
'    If xtype.Tag = "" Then
'        MsgBox "·„ Ì „  ÕœÌœ ‰Ê⁄ «·ÿ·»Ì…"
'        Exit Function
'    End If
'End If

If Not bIgShip Then
    If Trim(xShip_no.text) = "" Then
        MsgBox "—ﬁ„ »Ê·Ì’… «·‘Õ‰ „ÿ·Ê»"
        Exit Function
    End If
    
    If Not xship.MatchedWithList Then
        MsgBox "‘—ﬂ… «·‘Õ‰ €Ì— „”Ã·…"
        Exit Function
    End If
End If
'If IsDate(xdate1.caption) And xtype.Tag = "3" And Trim(xShip_No.text) = "" Then
'    MsgBox " «—ÌŒ «” ·«„ ›Ï „” ‰œ «” »œ«· »œÊ‰ —ﬁ„ »Ê·Ì’… ‘Õ‰"
'    Exit Function
'End If

If Not ValidQuant Then
    Exit Function
End If
'
'If Not ValidQuant(grid2) Then
'    Exit Function
'End If
myValid = True
End Function
Private Sub myload()
xdoc_no.Caption = CardTable!doc_no
xDate.text = myFormat_p(CardTable!Date)
xdate1.Caption = myFormat_p(CardTable!Date1)
xDate_mail.text = myFormat_p(CardTable!DATE_MAIL)
xtype.Tag = CardTable!Type
xtype.Caption = CardTable!TYPE_dESCA
xDiscount.text = CardTable!discount & ""
xShip_no.text = CardTable!ship_no & ""
xship.BoundText = CardTable!SHIP & ""
xdoc_no_sales.Caption = CardTable!doc_no_sales & ""
xSales_Replace.Caption = CardTable!sales_replace & ""
xValue.text = CardTable!Value & ""
xSales_ret.Caption = CardTable!SALES_RET & ""
xClosed.Value = IIf(CardTable!CLOSED, 1, 0)
xStage.Tag = CardTable!Stage & ""
xCharge2.text = CardTable!charge2 & ""
xStage.Caption = CardTable!stage_Desca & ""
xPayment_id.text = CardTable!PAYMENT_ID & ""
xNotes.text = CardTable!NOTES & ""
panel1(1).Caption = CardTable!USER_IP & ""

myLoadGrd

myloadGrdError


Handlecontrols LoadMode

CellPos 13, grid1.Rows - 2, grid1.Cols - 1
grdError.ShowCell grdError.Rows - 1, 1
Exit Sub
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub myDefine()
xdoc_no.Caption = ""
xDate.text = myFormat_p(Date)
xDate_mail.text = ""
xdoc_no_sales.Caption = ""
xSales_Replace.Caption = ""
xSales_ret.Caption = ""
xtype.Caption = xtype.TagVariant
xNotes.text = ""
xValue.text = ""
xtype.Tag = ""

xship.BoundText = ""
xShip_no.text = ""
xCharge2.text = ""
xValue.text = ""

addShip
bIg = True
xtotal_item.Caption = ""
xDiscount.text = ""
xtotal.Caption = ""
bIg = False
xClosed.Value = 0

grid1.Rows = 1
myAddItem
Fixgrd

grdError.Rows = 1
fixGrdError

Handlecontrols DefineMode
Fixgrd

CalcTotals
End Sub
Private Sub Handlecontrols(nMode)
bEditRecord = bEdit
bEditRecord = bEditRecord And cBranch = "00"
bEditRecord = bEditRecord And xtype.Tag <> ""
bEditRecord = bEditRecord And xClosed.Value = 0
bEditRecord = bEditRecord And xStage.Tag <> "2"


cmdSave.Enabled = bEditRecord
cmdSavePayment.Enabled = nMode = LoadMode And cmdSave.Enabled = False
cmdNewInv.Enabled = bEdit And cBranch = "00"

cmddel.Enabled = bEditRecord And nMode = LoadMode And cBranch = "00" And (Trim(xdate1.Caption) = "" Or xtype.Tag <> "3")

'cmdSaveDateMail.Enabled = xtype.Tag = "2" Or xtype.Tag = "3" Or xtype.Tag = "12"
xDate_mail.Enabled = xtype.Tag = "2" Or xtype.Tag = "3" Or xtype.Tag = "12"

cmdEdit.Enabled = xClosed.Value = 0 And xStage.Tag = 6

lblShip.Visible = xtype.Tag = "13"
xValue.Visible = xtype.Tag = "13"

xtype.Enabled = nMode = DefineMode

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

fmReplace.Visible = nMode = LoadMode And (xtype.Tag = "1" Or xtype.Tag = "11")
fmSend.Visible = nMode = LoadMode And (xtype.Tag = "2" Or xtype.Tag = "3" Or xtype.Tag = "11")

'lblReplace.Visible = nMode = LoadMode And xSales_replace.Caption <> ""
'xSales_replace.Visible = nMode = LoadMode And xSales_replace.Caption = ""

xdoc_no.Tag = nMode
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If Not bEditRecord Then
    Exit Sub
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
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
OpenCon con

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
grid1.TextMatrix(Row, 10) = "1"

'grid1.TextMatrix(i, 14) = fnBalance(grid1.TextMatrix(Row, 1), con, sStore)
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
    If Round(Val(xRate.text), nRound) <> Round(Val(xDiscount.text) / Val(xtotal_item.Caption) * 100, nRound) Then
        xRate.text = Myvalue(Round((Val(xDiscount.text) / Val(xtotal_item.Caption)) * 100, nRound))
    End If
Else
    xRate.text = ""
End If

xtotal.Caption = mRound(nTotalItem - Val(xDiscount.text), 2)
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
.ColWidth(3) = 1500
.ColWidth(4) = 0
.ColWidth(5) = 1500
.ColWidth(6) = 4000
.ColWidth(7) = 800
.ColWidth(8) = 1500
.ColWidth(9) = 1000
.ColWidth(10) = 800
.ColWidth(11) = 1000
.ColWidth(12) = 600
.ColWidth(13) = 1200
.ColWidth(16) = 1700
'.ColHidden(12) = True
.ColHidden(14) = True
.ColHidden(15) = True
.ColHidden(17) = True
.ColHidden(18) = True
.ColHidden(19) = True
.ColHidden(20) = True

.ColEditMask(10) = "999"
.ColHidden(.Cols - 2) = True
.ColHidden(.Cols - 1) = True
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
.Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = flexAlignRightCenter

Dim con As New ADODB.Connection
sOk = OpenCon(con)
Dim bGroup As Boolean
Dim sGroup As String
For i = 1 To grid1.Rows - 2
    If .ValueMatrix(i, 9) <> .ValueMatrix(i, 11) Then
        .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = vbRed
        .Cell(flexcpFontUnderline, i, 1, i, .Cols - 1) = vbRed
    End If
'    If sOk = "ok" And IsDgt(grid1.TextMatrix(i, 1)) Then
'        grid1.TextMatrix(i, 14) = fnBalance(grid1.TextMatrix(i, 1), con, sStore, , grid1.TextMatrix(grid1.Row, grid1.Cols - 1))
'    End If
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
    OpenCon con
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
OpenCon con

Dim aInsert As Variant
With oSalesRefund
Dim sdoc_no_new As String

sdoc_no_new = IncData("FILE6_90BH", "DOC_NO", con, , True)
If sdoc_no_new = "" Then sdoc_no_new = "1"

aInsert = AddFlag(Empty, "DOC_NO", sdoc_no_new)
aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(.xdiscount_ret.Caption))
aInsert = AddFlag(aInsert, "[ORDER_NO]", addstring(sOrder_no))
aInsert = AddFlag(aInsert, "[DATE]", addDate(Date))
aInsert = AddFlag(aInsert, "[TYPE]", oSalesRefund.sFlag)
aInsert = AddFlag(aInsert, "[SHIP]", addstring(xship.BoundText))
aInsert = AddFlag(aInsert, "[MAN]", addstring(xMan.Tag))
aInsert = AddFlag(aInsert, "[SHIP_NO]", addstring(xShip_no.text))
aInsert = AddFlag(aInsert, "[USERNAME]", addstring(cUserName))
aInsert = AddFlag(aInsert, "[USER_IP]", addstring(GetComputerName))
If oSalesRefund.sFlag = "2" Or oSalesRefund.sFlag = "3" Then
    aInsert = AddFlag(aInsert, "[SALES_RET]", addstring(.xdoc_no.Caption))
ElseIf oSalesRefund.sFlag = "11" Then
    aInsert = AddFlag(aInsert, "[SALES_REPLACE]", addstring(.xdoc_no.Caption))
End If
aInsert = AddFlag(aInsert, "[STAGE]", "7")
End With

con.BeginTrans
On Error GoTo myerror
con.Execute addInsert(aInsert, "FILE6_90BH")

With oSalesRefund.grid1
For i = 1 To .Rows - 2
    aInsert = AddFlag(Empty, "ITEM", addstring(.TextMatrix(i, 1)))
    aInsert = AddFlag(aInsert, "QUANT", .ValueMatrix(i, 8))
    aInsert = AddFlag(aInsert, "PRICE", .ValueMatrix(i, 9))
    aInsert = AddFlag(aInsert, "PRICE_C", .ValueMatrix(i, 7))
    aInsert = AddFlag(aInsert, "COST", .ValueMatrix(i, 15))
    aInsert = AddFlag(aInsert, "MAN", addstring(.TextMatrix(i, 15 + 1)))
    aInsert = AddFlag(aInsert, "USER_IP", addstring(GetComputerName))
    aInsert = AddFlag(aInsert, "OFFER_NO", addstring(.TextMatrix(i, 11)))
    aInsert = AddFlag(aInsert, "DOC_OFFER", addstring(.TextMatrix(i, 12)))
    aInsert = AddFlag(aInsert, "DISCOUNT_OFFER", .ValueMatrix(i, 13))
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
Public Function myReplaceRefundAll(pDoc_no As String) As Boolean

dbm.OpenCon
Dim locTable As New ADODB.Recordset
Set locTable = dbm.myRs("SELECT * FROM FILE6_20H WHERE DOC_NO = " & MyParn(pDoc_no))

Dim aInsert As Variant
Dim sdoc_no_new As String
sdoc_no_new = dbm.IncData("FILE6_90BH", "DOC_NO", , True)
If sdoc_no_new = "" Then sdoc_no_new = "1"
aInsert = AddFlag(Empty, "DOC_NO", addstring(sdoc_no_new))
aInsert = AddFlag(aInsert, "[DISCOUNT]", locTable!discount)
aInsert = AddFlag(aInsert, "[ORDER_NO]", addstring(sOrder_no))
aInsert = AddFlag(aInsert, "[DATE]", addDate(Date))
aInsert = AddFlag(aInsert, "[TYPE]", "3")
aInsert = AddFlag(aInsert, "[MAN]", addstring(locTable!MAN))
aInsert = AddFlag(aInsert, "[SHIP]", addstring(locTable!SHIP))
aInsert = AddFlag(aInsert, "[SHIP_NO]", addstring(locTable!ship_no))
aInsert = AddFlag(aInsert, "CHARGE2", Val(locTable!charge2 & ""))
aInsert = AddFlag(aInsert, "[USERNAME]", addstring(cUserName))
aInsert = AddFlag(aInsert, "[USER_IP]", addstring(GetComputerName))
aInsert = AddFlag(aInsert, "[SALES_RET]", addstring(pDoc_no))
aInsert = AddFlag(aInsert, "[STAGE]", "7")

'con.BeginTrans
On Error GoTo myerror
dbm.addSql addInsert(aInsert, "FILE6_90BH")

Set locTable = New ADODB.Recordset
Set locTable = dbm.myRs("SELECT * FROM FILE6_20 WHERE DOC_NO = " & MyParn(pDoc_no))

Do Until locTable.EOF
    aInsert = AddFlag(Empty, "ITEM", locTable!Item)
    aInsert = AddFlag(aInsert, "QUANT", locTable!Quant)
    aInsert = AddFlag(aInsert, "PRICE", locTable!price)
    aInsert = AddFlag(aInsert, "PRICE_C", locTable!price_c)
    aInsert = AddFlag(aInsert, "COST", locTable!cost)
    aInsert = AddFlag(aInsert, "USER_IP", addstring(GetComputerName))
    aInsert = AddFlag(aInsert, "OFFER_NO", addstring(locTable!OFFER_NO))
    aInsert = AddFlag(aInsert, "MAN", addstring(locTable!MAN))
    aInsert = AddFlag(aInsert, "DOC_OFFER", addstring(locTable!DOC_OFFER))
    aInsert = AddFlag(aInsert, "DISCOUNT_OFFER", locTable!DISCOUNT_OFFER)
    aInsert = AddFlag(aInsert, "DOC_NO", addstring(sdoc_no_new))
    dbm.addSql addInsert(aInsert, "FILE6_90B")
    locTable.MoveNext
Loop

If Not dbm.ExecuteTransaction Then
    dbm.closeCon
    Exit Function
End If

dbm.closeCon
Set locTable = Nothing

myReplaceRefundAll = True

MsgBox " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"

If Not openCardTable(tbMode.tbFind, sdoc_no_new) Then
    If Not openCardTable Then
        myDefine
    End If
End If
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function

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
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If bStopCell Then
    bStopCell = False
ElseIf Col = 1 And grid1.TextMatrix(Row, Col) = "" Then
ElseIf Col < 10 Then
    grid1.Col = 10
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

'On Error GoTo myerror
cFilter = retFilter

Set CardTable = New ADODB.Recordset
cString = "SELECT TOP 1 FILE6_90BH.*," & _
          " FILE6_90BH.INV_NO AS DOC_NO_SALES," & _
          " ONLINE_TYPE_CODES.DESCA AS TYPE_DESCA," & _
          " s.DESCA AS STAGE_DESCA" & _
          " FROM FILE6_90BH " & _
          " INNER JOIN ONLINE_TYPE_CODES ON FILE6_90BH.[TYPE] = ONLINE_TYPE_CODES.CODE" & _
          " LEFT JOIN STAGES_CODES s ON FILE6_90BH.[STAGE] = s.CODE"
   
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
Private Function ValidQuant() As Boolean
With grid1
For i = 1 To .Rows - 2
    If .ValueMatrix(i, 10) <= 0 Then
        Exit Function
    End If
Next
End With
ValidQuant = True
End Function
Private Function retRecords(pDoc_no, ByRef nRecords As Long, ByRef nRecord As Long) As Variant
Dim cString As String, locTable As New ADODB.Recordset
If pDoc_no <> "" Then
    cString = "SELECT Count(FILE6_90BH.DOC_NO) AS records,COUNT(CASE WHEN FILE6_90BH.DOC_NO <= " & MyParn(pDoc_no) & " THEN 1 END) AS record"
Else
    cString = "SELECT Count(FILE6_90BH.DOC_NO) AS records,0 as record"
End If

cString = cString & " FROM FILE6_90BH"

If cFilter <> "" Then
    cString = cString & " WHERE " & cFilter
End If

Set locTable = myRs(cString)
If Not locTable.EOF Then
    nRecords = locTable!RECORDS
    nRecord = Val(locTable!Record & "")
End If
End Function
Private Function retFilter()
retFilter = "FILE6_90BH.ORDER_NO = " & MyParn(sOrder_no)
If sStage <> "" Then
    retFilter = retFilter & " AND FILE6_90BH.STAGE = " & sStage
End If
If bOpenOnly Then
    retFilter = retFilter & " AND FILE6_90BH.CLOSED = 0"
End If
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

Private Sub SSCommand1_Click()
End Sub

Private Sub xdate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xDate_LostFocus()
myLostFocus xDate
myValidDate xDate
End Sub

Private Sub xdate_mail_GotFocus()
myGotFocus xDate_mail
End Sub
Private Sub xdate_mail_LostFocus()
myLostFocus xDate_mail
myValidDate xDate_mail
End Sub

Private Sub xDiscount_Change()
If bIg Then Exit Sub
CalcTotals
End Sub

Private Sub xPayment_id_GotFocus()
myGotFocus xPayment_id
End Sub
Private Sub xPayment_id_LostFocus()
myLostFocus xPayment_id
End Sub
Private Sub Notes_GotFocus()
myGotFocus NOTES
End Sub
Private Sub Notes_LostFocus()
myLostFocus NOTES
End Sub

Private Sub HandleInit()
If cBranch = "00" Then
    nUser = enUser.MainUser
Else
    nUser = enUser.branch
End If
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
       
Dim locTable As New ADODB.Recordset
Set locTable = myRs(cString)

If locTable.EOF And locTable.BOF Then
    MsgBox "«·„” ‰œ €Ì— „”Ã·"
    Exit Function
End If

If IsNull(locTable!MAN) Then
    sMan = "0001"
Else
    sMan = locTable!MAN
End If

aInsert = AddFlag(Empty, "CODE", addstring("0000"))
aInsert = AddFlag(aInsert, "ONLINE", "2")
aInsert = AddFlag(aInsert, "[Date]", addDate(sDate))
aInsert = AddFlag(aInsert, "STORE", addstring(cBranchStore))
aInsert = AddFlag(aInsert, "BOX", addstring(cManBox))
aInsert = AddFlag(aInsert, "DISCOUNT", Val(locTable!discount))
If Val(locTable!discount & "") <> 0 Then
    aInsert = AddFlag(aInsert, "username_disc", addstring("«Ê‰·«Ì‰"))
End If

aInsert = AddFlag(aInsert, "BRANCH", addstring(cBranch))
aInsert = AddFlag(aInsert, "PHONE", addstring(locTable!Phone))
aInsert = AddFlag(aInsert, "MAN", addstring(locTable!MAN))

aInsert = AddFlag(aInsert, "INV_TYPE_ONLINE", locTable!Type)
aInsert = AddFlag(aInsert, "DOC_NO_ONLINE", locTable!doc_no)
aInsert = AddFlag(aInsert, "ONLINE_DOC", addstring(locTable!ORDER_NO))
aInsert = AddFlag(aInsert, "SALES_RET", addstring(locTable!SALES_RET))
    
aInsert = AddFlag(aInsert, "PAYMENT_ID", addstring(locTable!PAYMENT_ID))
aInsert = AddFlag(aInsert, "userName", addstring(cUserName))
aInsert = AddFlag(aInsert, "USER_IP", addstring(GetComputerName))

cAddress = locTable!Shipping_City & ""
If locTable!Street & "" <> "" Then
    cAddress = cAddress & Tr(cAddress, " ") & locTable!Street
End If
aInsert = AddFlag(aInsert, "ADDRESS", addstring(cAddress))
aInsert = AddFlag(aInsert, "[NAME]", addstring(locTable!Name))
aInsert = AddFlag(aInsert, "[city]", addstring(locTable!city))
aInsert = AddFlag(aInsert, "[Shipping_City]", addstring(locTable!Shipping_City))
aInsert = AddFlag(aInsert, "[street]", addstring(locTable!Street))
aInsert = AddFlag(aInsert, "[Payment_Method]", addstring(locTable!Payment_Method))

sDoc_New = NewflagDocRs(sDate, cManBox)
aInsert = AddFlag(aInsert, "[doc_no]", addstring(sDoc_New))

Dim con As New ADODB.Connection
If OpenCon(con) <> "ok" Then
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
Private Sub xDate1_GotFocus()
myGotFocus xdate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xdate1
myValidDate xdate1
End Sub

Private Sub xRate_Change()
If bIg Then Exit Sub
bIg = True
If Val(xtotal_item.Caption) <> 0 Then
    If Round(Val(xRate.text), 0) <> Round(Val(xDiscount.text) / Val(xtotal_item.Caption) * 100, 0) Then
        xDiscount.text = Round((Val(xRate.text) * Val(xtotal_item.Caption)) / 100, 2)
    End If
Else
    xDiscount.text = ""
End If
CalcTotals
bIg = False
End Sub

Private Sub xSales_replace_Click()
ItemsLook Me, oSearchItem, , , "FILE1_10.ITEM IN(SELECT ITEM FROM FILE6_20 WHERE DOC_NO = " & MyParn(xSales_Replace.Caption) & ")", False
End Sub

Private Sub myloadGrdError()
Dim strSql As String
strSql = "SELECT FORMAT([TIME], 'yyyy/M/d HH:mm') AS [«·Êﬁ ]" & _
          ",[BRANCH] AS [«·„Œ“‰]" & _
          ",[MSG] AS [«·—”«·…]" & _
          ",[DESCA_MAN] AS [«·„—”·]" & _
          ",[STAGE_DESCA] AS [«·„—Õ·…]" & _
          ",[DOC_NO_SUP]" & _
          ",[ORDER_NO]" & _
          ",[TYPE]" & _
           ",[STAGE]" & _
          " From [dbo].[vw_msg]" & _
          " WHERE ORDER_NO = " & MyParn(sOrder_no) & _
          " AND DOC_NO_SUP = " & xdoc_no.Caption & _
          " ORDER BY TIME"

On Error GoTo myerror

With grdError
Set .DataSource = myRs(strSql)

fixGrdError
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub fixGrdError()
With grdError
.TextMatrix(i, 0) = "«·Êﬁ "
.TextMatrix(i, 1) = "«·„Œ“‰"
.TextMatrix(i, 2) = "«·—”«·…"
.TextMatrix(i, 3) = "«·„—”·"
.TextMatrix(i, 4) = "«·„—Õ·…"

.ColWidth(0) = 1800
.ColWidth(1) = 1300
.ColWidth(2) = 8000
.ColWidth(3) = 1800
.ColWidth(4) = 1800

.ColHidden(.Cols - 4) = True
.ColHidden(.Cols - 3) = True
.ColHidden(.Cols - 2) = True
.ColHidden(.Cols - 1) = True

For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
grdError.ShowCell grdError.Rows - 1, 1
End With
End Sub
Private Sub xtype_Click()
On Error GoTo myerror
OnlineTypeLookup Me, oSearchType, filterType, , , "type"
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Function filterType() As String
Dim cString As String
cString = "SELECT " & _
          "COUNT(CASE WHEN TYPE = 1 THEN 1 END) as TYPE1," & _
          "COUNT(CASE WHEN TYPE = 2 THEN 1 END) as TYPE2," & _
          "COUNT(CASE WHEN TYPE = 3 THEN 1 END) as TYPE3," & _
          "COUNT(CASE WHEN TYPE = 11 THEN 1 END) as TYPE11," & _
          "COUNT(CASE WHEN TYPE = 12 THEN 1 END) as TYPE12" & _
          " FROM FILE6_90BH " & _
          " WHERE FILE6_90BH.ORDER_NO = " & MyParn(sOrder_no) & _
          " AND FILE6_90BH.CLOSED = 0"

Dim locTable As New ADODB.Recordset
Set locTable = myRs(cString)
If Not locTable.EOF Then
    If locTable!type1 + locTable!type3 + locTable!type11 + locTable!type12 = 0 And locTable!type2 > 0 Then filterType = "1"
    If locTable!type2 + locTable!type3 + locTable!type11 + locTable!type12 = 0 Then filterType = filterType & Tr(filterType, ",") & "2"
    If locTable!type1 + locTable!type2 + locTable!type3 + locTable!type11 + locTable!type12 = 0 Then filterType = filterType & Tr(filterType, ",") & "3"
    If locTable!type1 + locTable!type2 + locTable!type3 + locTable!type11 = 0 Then filterType = filterType & Tr(filterType, ",") & "11"
    If locTable!type1 + locTable!type2 + locTable!type3 + locTable!type12 = 0 Then filterType = filterType & Tr(filterType, ",") & "12"
End If
If filterType = "" Then
    filterType = "(1 = 2)"
Else
    filterType = "CODE IN(" & filterType & ",13)"
End If
Set locTable = Nothing
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
Private Sub addShip()
Dim locTable As New ADODB.Recordset
On Error GoTo myerror
Set locTable = myRs("SELECT TOP 1 * FROM FILE6_90BH WHERE ORDER_NO = " & MyParn(sOrder_no) & " AND  CLOSED = 0 AND (TYPE = 2 OR TYPE = 11 OR TYPE = 12)")
If Not locTable.EOF Then
    xShip_no.text = locTable!ship_no & ""
    xship.BoundText = locTable!SHIP & ""
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

