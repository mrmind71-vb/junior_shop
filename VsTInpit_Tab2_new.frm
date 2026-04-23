VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VsTInpit_Tab 
   Caption         =   " ﬁ«—Ì— «·Ê«—œ „Ã„⁄…"
   ClientHeight    =   10230
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   18585
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10230
   ScaleWidth      =   18585
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdFixPrice 
      Caption         =   "÷»ÿ  ﬂ·›…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3915
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   405
      Width           =   2175
   End
   Begin VB.Frame Frame7 
      Height          =   1320
      Left            =   3915
      RightToLeft     =   -1  'True
      TabIndex        =   62
      Top             =   945
      Width           =   2175
      Begin VB.OptionButton optPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "”⁄— „” Â·ﬂ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   675
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   945
         Width           =   1365
      End
      Begin VB.OptionButton optPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "”⁄—  ﬂ·›…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   585
         Width           =   1320
      End
      Begin VB.OptionButton optPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "”⁄—  ﬂ·›… Ê„” Â·ﬂ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   225
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2220
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   11580
      Begin VB.TextBox XGRMODEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   135
         MaxLength       =   15
         TabIndex        =   38
         Top             =   270
         Width           =   1545
      End
      Begin VB.TextBox xModelFact 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   15
         TabIndex        =   36
         Top             =   270
         Width           =   1545
      End
      Begin VB.TextBox XMOSM2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   135
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   1710
         Width           =   825
      End
      Begin VB.ComboBox xmosm20 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1710
         Width           =   825
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   5670
         TabIndex        =   2
         Top             =   1395
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xFact 
         Height          =   315
         Left            =   6450
         TabIndex        =   3
         Top             =   1035
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   135
         TabIndex        =   4
         Top             =   630
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSupp 
         Height          =   315
         Left            =   6450
         TabIndex        =   5
         Top             =   690
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XSTORE 
         Height          =   315
         Left            =   135
         TabIndex        =   12
         Top             =   990
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777088
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XSECTION 
         Height          =   315
         Left            =   135
         TabIndex        =   14
         Top             =   1350
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker xdate1 
         Height          =   375
         Left            =   8145
         TabIndex        =   21
         Top             =   270
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   661
         _Version        =   393216
         MousePointer    =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   133038081
         CurrentDate     =   43097
      End
      Begin MSComCtl2.DTPicker xdate2 
         Height          =   375
         Left            =   6300
         TabIndex        =   22
         Top             =   270
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   661
         _Version        =   393216
         MousePointer    =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   133038081
         CurrentDate     =   43097
      End
      Begin MSDataListLib.DataCombo xage 
         Height          =   315
         Left            =   7935
         TabIndex        =   28
         Top             =   1755
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xsex 
         Height          =   315
         Left            =   5670
         TabIndex        =   29
         Top             =   1755
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "«·›—⁄"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   1080
         Width           =   1230
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ „Ã„Ê⁄… "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1725
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   315
         Width           =   1005
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ „ÊœÌ·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4515
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   315
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "„Ê”„ œ«Œ·Ï"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   4590
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1755
         Width           =   1050
      End
      Begin VB.Label Label2 
         Caption         =   "«·‰Ê⁄ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   7245
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   1800
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "›∆… ⁄„—Ì… :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   9990
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "‰Ê⁄ «·„Ê”„ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1755
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "«·ﬁ”„ "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1395
         Width           =   600
      End
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1080
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·› —… :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9990
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label2 
         Caption         =   "„Ê”„ «·„ÊœÌ·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   720
         Width           =   1230
      End
      Begin VB.Label Label3 
         Caption         =   "«·„’‰⁄ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9990
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1065
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "„Ã„Ê⁄… „ÊœÌ·«  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   9990
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1440
         Width           =   1410
      End
      Begin VB.Label Label5 
         Caption         =   "«·„Ê—œ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9990
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   675
         Width           =   1545
      End
   End
   Begin VB.Frame Frame6 
      Height          =   690
      Left            =   11925
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   2205
      Width           =   5775
      Begin VB.CheckBox chkGroup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·„Ã„Ê⁄…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   270
         Width           =   1140
      End
      Begin VB.CheckBox chkMosm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·„Ê”„"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   270
         Width           =   915
      End
      Begin VB.CheckBox chkAge 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·›∆… «·⁄„—Ì…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   270
         Width           =   1275
      End
      Begin VB.CheckBox chkSection 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·ﬁ”„"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   270
         Width           =   780
      End
      Begin VB.CheckBox chkType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·‰Ê⁄"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   270
         Width           =   780
      End
   End
   Begin VB.Frame Frame3 
      Height          =   690
      Left            =   4005
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   2250
      Width           =   2085
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "query "
         Height          =   285
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   225
         Width           =   825
      End
      Begin VB.CheckBox XOutlet 
         Alignment       =   1  'Right Justify
         Caption         =   "Outlet "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   225
         Width           =   870
      End
   End
   Begin VB.Frame fmBranch 
      Height          =   690
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   2205
      Width           =   2085
      Begin VB.CheckBox chkStoped 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "≈ŸÂ«— ›—Ê⁄ „ Êﬁ›…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   225
         Width           =   1860
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   2190
      Left            =   90
      TabIndex        =   46
      Top             =   90
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   3863
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "’Ê—… «·„ÊœÌ·"
      TabPicture(0)   =   "VsTInpit_Tab2_new.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ImgX1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "query "
      TabPicture(1)   =   "VsTInpit_Tab2_new.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "textquery"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.TextBox textquery 
         Alignment       =   1  'Right Justify
         Height          =   1770
         Left            =   45
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   360
         Width           =   3660
      End
      Begin ImgXCtrl6.ImgXCtrl ImgX1 
         Height          =   1770
         Left            =   -74910
         TabIndex        =   47
         Top             =   360
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   3122
         BackColor       =   16777215
         BorderStyle     =   4
         AutoZoom        =   -1  'True
         SelectionLineType=   4
         Center          =   -1  'True
         ImageBorderThickness=   1
         AutoZoomType    =   3
         DoubleBuffer    =   -1  'True
         LicenseUserName =   "mrmind"
         LicenseRegCode  =   "íß“ªª•≤≥Ω≠∞“±≤ß´¥©ÆØOOHH-FAOOYNJB-EQCF6gI"
      End
   End
   Begin VB.Frame Frame5 
      Height          =   690
      Left            =   -900
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   4095
      Visible         =   0   'False
      Width           =   1635
      Begin MSAdodcLib.Adodc DATA10 
         Height          =   330
         Left            =   -540
         Top             =   585
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   582
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
      Begin MSAdodcLib.Adodc DATA11 
         Height          =   330
         Left            =   -405
         Top             =   495
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   582
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
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   -315
         Top             =   540
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   582
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
      Begin MSAdodcLib.Adodc DATA14 
         Height          =   330
         Left            =   -180
         Top             =   540
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   582
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
      Begin MSAdodcLib.Adodc DATA15 
         Height          =   330
         Left            =   0
         Top             =   450
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   582
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
      Begin MSAdodcLib.Adodc DATA16 
         Height          =   330
         Left            =   -45
         Top             =   450
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   582
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
      Begin MSAdodcLib.Adodc DATA17 
         Height          =   330
         Left            =   0
         Top             =   540
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   582
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
   End
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2250
      Width           =   3885
      Begin VB.CommandButton cmdExel 
         Height          =   510
         Left            =   1035
         Picture         =   "VsTInpit_Tab2_new.frx":0038
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   915
      End
      Begin VB.CommandButton cmdGo 
         Height          =   510
         Left            =   2925
         Picture         =   "VsTInpit_Tab2_new.frx":2823
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   915
      End
      Begin VB.CommandButton cmdExit 
         Height          =   510
         Left            =   45
         Picture         =   "VsTInpit_Tab2_new.frx":4D15
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   135
         Width           =   960
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   510
         Left            =   1980
         Picture         =   "VsTInpit_Tab2_new.frx":7181
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   135
         Width           =   915
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   9900
      Width           =   18585
      _ExtentX        =   32782
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin MSAdodcLib.Adodc data5 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin MSAdodcLib.Adodc DATA6 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin MSAdodcLib.Adodc DATA7 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Caption         =   "data7"
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
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   32
      Top             =   9705
      Visible         =   0   'False
      Width           =   18585
      _ExtentX        =   32782
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Caption         =   "DATA12"
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
   Begin MSAdodcLib.Adodc DATA13 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin MSAdodcLib.Adodc DATA12 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6675
      Left            =   90
      TabIndex        =   40
      Top             =   3015
      Width           =   17610
      _ExtentX        =   31062
      _ExtentY        =   11774
      _Version        =   393216
      Tabs            =   8
      Tab             =   2
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "„ÊœÌ·« "
      TabPicture(0)   =   "VsTInpit_Tab2_new.frx":95AB
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grid0"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "„’«‰⁄"
      TabPicture(1)   =   "VsTInpit_Tab2_new.frx":95C7
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GRID1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   " «·√ﬁ”«„"
      TabPicture(2)   =   "VsTInpit_Tab2_new.frx":95E3
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "GRID2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "«·„Ã„Ê⁄« "
      TabPicture(3)   =   "VsTInpit_Tab2_new.frx":95FF
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "GRID3"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   " ›∆«  ⁄„—Ì…"
      TabPicture(4)   =   "VsTInpit_Tab2_new.frx":961B
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "GRID4"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   " ›’Ì·Ï „Ã„⁄ - „ÊœÌ·«  - —»Õ"
      TabPicture(5)   =   "VsTInpit_Tab2_new.frx":9637
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "grid5"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "«·›—Ê⁄"
      TabPicture(6)   =   "VsTInpit_Tab2_new.frx":9653
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "GRID6"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "„Ã„Ê⁄«  „ÊœÌ·« "
      TabPicture(7)   =   "VsTInpit_Tab2_new.frx":966F
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "GRID7"
      Tab(7).ControlCount=   1
      Begin VSFlex7Ctl.VSFlexGrid GRID1 
         Height          =   6165
         Left            =   -74955
         TabIndex        =   41
         Top             =   360
         Width           =   17250
         _cx             =   30427
         _cy             =   10874
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16776960
         ForeColorSel    =   64
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   21
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
         WordWrap        =   0   'False
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
      Begin VSFlex7Ctl.VSFlexGrid GRID2 
         Height          =   6165
         Left            =   45
         TabIndex        =   42
         Top             =   360
         Width           =   17475
         _cx             =   30824
         _cy             =   10874
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16776960
         ForeColorSel    =   64
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   19
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
         WordWrap        =   0   'False
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
      Begin VSFlex7Ctl.VSFlexGrid GRID3 
         Height          =   6255
         Left            =   -74955
         TabIndex        =   43
         Top             =   360
         Width           =   17475
         _cx             =   30824
         _cy             =   11033
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16776960
         ForeColorSel    =   64
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   21
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
         WordWrap        =   0   'False
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
      Begin VSFlex7Ctl.VSFlexGrid GRID4 
         Height          =   6255
         Left            =   -74955
         TabIndex        =   59
         Top             =   360
         Width           =   17475
         _cx             =   30824
         _cy             =   11033
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16776960
         ForeColorSel    =   64
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   21
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
         WordWrap        =   0   'False
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
      Begin VSFlex7Ctl.VSFlexGrid GRID6 
         Height          =   6255
         Left            =   -74955
         TabIndex        =   60
         Top             =   360
         Width           =   17475
         _cx             =   30824
         _cy             =   11033
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16776960
         ForeColorSel    =   64
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   21
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
         WordWrap        =   0   'False
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
      Begin VSFlex7Ctl.VSFlexGrid GRID7 
         Height          =   6255
         Left            =   -74955
         TabIndex        =   61
         Top             =   360
         Width           =   17520
         _cx             =   30903
         _cy             =   11033
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16776960
         ForeColorSel    =   64
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   21
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
         WordWrap        =   0   'False
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
      Begin VSFlex7Ctl.VSFlexGrid grid0 
         Height          =   6165
         Left            =   -74910
         TabIndex        =   67
         Top             =   360
         Width           =   17430
         _cx             =   30745
         _cy             =   10874
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16776960
         ForeColorSel    =   64
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   23
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
         WordWrap        =   0   'False
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
      Begin VSFlex7Ctl.VSFlexGrid grid5 
         Height          =   6255
         Left            =   -74955
         TabIndex        =   68
         Top             =   360
         Width           =   17475
         _cx             =   30824
         _cy             =   11033
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16776960
         ForeColorSel    =   64
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   28
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
         WordWrap        =   0   'False
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
   End
   Begin MSAdodcLib.Adodc DATA8 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   582
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
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   8235
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   2205
      Width           =   3660
      Begin VB.OptionButton xall 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "›—Ê⁄ ›ﬁÿ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   270
         Width           =   1050
      End
      Begin VB.OptionButton xall 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   " ÊﬂÌ·«  ›ﬁÿ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   1035
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   315
         Width           =   1230
      End
      Begin VB.OptionButton xall 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·ﬂ·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   315
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Label xModelNo 
      Alignment       =   1  'Right Justify
      Caption         =   "Label6"
      Height          =   285
      Left            =   1125
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2520
      Visible         =   0   'False
      Width           =   2310
   End
End
Attribute VB_Name = "VsTInpit_Tab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oSearchItem As New Search3
Dim oSearch As New Search3
Dim bCost As Boolean
Dim bPrice As Boolean
Dim con As New ADODB.Connection
Private Sub CMD_PRINT_Click()
End Sub
Private Sub cmdExel_Click()
    Select Case SSTab1.Tab
        Case 0
            ToFileExel2 grid0, , , , , 1.1, , , , , , Me
        Case 1
            ToFileExel2 grid1, , , , , 1.1, , , , , , Me
        Case 2
            ToFileExel2 GRID2, , , , , 1.1, , , , , , Me
        Case 3
            ToFileExel2 GRID3, , , , , 1.1, , , , , , Me
        Case 4
            ToFileExel2 GRID4, , , , , 1.1, , , , , , Me
        Case 5
            ToFileExel2 GRID5, , , , , 1.1, , , , , , Me
        Case 6
            ToFileExel2 GRID6, , , , , 1.1, , , , , , Me
        Case 7
            ToFileExel2 GRID7, , , , , 1.1, , , , , , Me
    End Select
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFixPrice_Click()
PriceSalesFix.Show 1
End Sub

Private Sub cmdGo_Click()
   
    Select Case SSTab1.Tab
        Case 0
            myload0
        Case 1
            MYLOAD1
        Case 2
            MYLOAD2
        Case 3
            MYLOAD3
        Case 4
            MYLOAD4
        Case 5
            myload5
        Case 6
            myload6
        Case 7
            myload7
    End Select
    
    
'    myload0
'    MYLOAD1
'    MYLOAD2
'
'    MYLOAD3
'    myload4
'    myload5
'    myload6
'    myload7
End Sub
Private Sub cmdPrint_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    Dim cHead3 As String
    cHead1 = Me.Caption & xStore.text & " " & xsex.text & " " & xage.text & "  " & xModelFact.text
    cHead2 = " „‰  «—ÌŒ " & Format(xDate1.Value, "DD-MM-YYYY") & " Õ Ï  «—ÌŒ " & Format(xDate2.Value, "DD-MM-YYYY")
    cHead3 = ""
    If xFact.MatchedWithList Then cHead3 = cHead3 & " „’‰⁄ " & xFact.text
    If xSupp.MatchedWithList Then cHead3 = cHead3 & " „Ê—œ " & xSupp.text
    If xGroup.MatchedWithList Then cHead3 = cHead3 & " „Ã„Ê⁄…" & xGroup.text
    If XSECTION.text <> "" Then cHead3 = cHead3 & " ﬁ”„ " & XSECTION.text
    If xMosm.MatchedWithList Then cHead3 = cHead3 & " „Ê”„ " & xMosm.text
    printGrdNew.DOPRINT grid1, 0.8, -2, cHead1, cHead2, , , , False, 9
    printGrdNew.Show 1
End Sub
Private Sub Form_Load()
    openCon con
    bCost = True
    bPrice = True
    
    textquery.Visible = bSupermode And lSupperVisor
    Check1.Visible = bSupermode And lSupperVisor
    
    xmosm20.AddItem " "
    xmosm20.AddItem "S"
    xmosm20.AddItem "W"
    xmosm20.AddItem "M"
    
    Set data8.Recordset = myRecordSet("Select Code,DescA From File0_82  order by Desca", con)
    Set xage.RowSource = data8
    xage.ListField = "Desca"
    xage.BoundColumn = "Code"
    
    Set DATA6.Recordset = myRecordSet("Select Code,DescA From File0_81  order by Desca", con)
    Set xsex.RowSource = DATA6
    xsex.ListField = "Desca"
    xsex.BoundColumn = "Code"
    
    Set data7.Recordset = myRecordSet("Select Code,DescA From STORE_ALL WHERE " & cUserStore, con)
    Set xStore.RowSource = data7
    xStore.ListField = "Desca"
    xStore.BoundColumn = "Code"
    If cBranch <> "00" Then
        xStore.BoundText = cBranchStore
        xStore.Enabled = False
    End If
    
    Set data1.Recordset = myRecordSet("Select Code,DescA From File1_10SC ORDER BY code ", con)
    Set XSECTION.RowSource = data1
    XSECTION.ListField = "Desca"
    XSECTION.BoundColumn = "Code"
    
    Set DATA2.Recordset = myRecordSet("Select Code,DescA From File4_10 order by Desca", con)
    Set xSupp.RowSource = DATA2
    xSupp.ListField = "Desca"
    xSupp.BoundColumn = "Code"
    
    Set DATA3.Recordset = myRecordSet("Select Code,DescA From File1_50 ORDER BY DESCA", con)
    Set xGroup.RowSource = DATA3
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    
    Set DATA4.Recordset = myRecordSet("Select mosm ,descA From mosm ORDER BY date DESC ", con)
    Set xMosm.RowSource = DATA4
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "MOSM"
    xMosm.BoundText = cPMosm
    
    
    Set DATA5.Recordset = myRecordSet("Select code ,desca From fact ORDER BY DESCA ", con)
    Set xFact.RowSource = DATA5
    xFact.ListField = "Desca"
    xFact.BoundColumn = "Code"
    
    Set grid0.DataSource = data10
    Set grid1.DataSource = DATA11
    Set GRID2.DataSource = data12
    Set GRID3.DataSource = DATA13
    Set GRID4.DataSource = DATA14
    Set GRID5.DataSource = DATA15
    Set GRID6.DataSource = DATA16
    Set GRID7.DataSource = DATA17

    FIXGRID0
    FixGrid1
    FixGrid2
    FixGrid3
    FIXGRID4
    FIXGRID5
    FIXGRID6
    FIXGRID7
    
    fmBranch.Visible = lSupperVisor
End Sub
Private Sub myload0()
Dim cString  As String
Dim cFields As String, cWhere As String, cWhereField As String
With grid0
    If IsDate(xDate1.Value) Then
        cWhereField = " AND DATE >= " & DateSq(xDate1.Value)
    End If
    
    ' Ê«—œ
    cFields = cFields & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]) ")
    ' Ê«—œ »”⁄— «· ﬂ·›…
    
    If bCost Then
        cFields = cFields & "," & _
                myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "(FILE1_11_ALL.[IN] * FILE1_11_ALL.COST)")
    End If
    
    ' Ê«—œ »”⁄— «·„” Â·ﬂ
    If bPrice Then
        cFields = cFields & "," & _
                myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "(FILE1_11_ALL.[IN] * FILE1_11_ALL.PRICE_SALES)")
    End If
    
    ' ’«œ—
    cFields = cFields & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([OUT])")
    
    ' ’«œ— »”⁄— «· ﬂ·›…
    If bCost Then
        cFields = cFields & "," & _
                myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "(FILE1_11_ALL.[OUT] * FILE1_11_ALL.COST)")
    End If
    
    ' ’«œ— »”⁄— «·„” Â·ﬂ
    If bPrice Then
        cFields = cFields & "," & _
                myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "(FILE1_11_ALL.[OUT] * FILE1_11_ALL.PRICE_SALES)")
    End If
    
    ' ’«›Ì «·ﬂ„Ì…
        
    cFields = cFields & "," & _
              myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) ")
    
    ' ’«›Ì »”⁄— «· ﬂ·›…
    If bCost Then
        cFields = cFields & "," & _
                myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "(([IN]-[OUT]) * FILE1_11_ALL.COST)")
    End If

    '  ’«›Ì »”⁄— «·„” Â·ﬂ
    If bPrice Then
        cFields = cFields & "," & _
                myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "(([IN]-[OUT]) * FILE1_11_ALL.PRICE_SALES)")
    End If
    
    
    cFields = cFields & ",0," & _
            myiif("(type = '6')" & cWhereField, "([OUT])")

    If bCost Then
        cFields = cFields & "," & _
                myiif("(type = '6')" & cWhereField, "([OUT] * FILE1_11_ALL.COST)")
    End If
    
    If bPrice Then
        cFields = cFields & "," & _
                myiif("(type = '6')" & cWhereField, "([OUT] * FILE1_11_ALL.PRICE_SALES)")
    End If
                    
    cFields = cFields & "," & _
              "SUM([IN]-[OUT])"
    
    If bCost Then
        cFields = cFields & "," & _
                 " ROUND(SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ),2)"
    End If
    
    If bPrice Then
        cFields = cFields & "," & _
                " ROUND(SUM(([IN]-[OUT]) * FILE1_11_ALL.PRICE_SALES ),2)"
    End If
                    
    cString = " SELECT  FILE4_10.DESCA AS supldesca, FACT.DESCA AS factdesca, FILE1_10.MOSM, FILE1_10.MODELFACT0, FILE1_10.desca " & _
                 cFields & ", sum(([in]- [out]) * FILE1_10.COSTitem) , FILE1_10.MODELNO " & _
                " FROM  FILE1_10 INNER JOIN FILE4_10 ON FILE1_10.code = FILE4_10.CODE" & _
                " INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE" & _
                " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM " & _
                " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
                " WHERE file1_10.ISNOITEM = 0 "
    
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    If xMosm.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING (FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    
    If xmosm20.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm20.text)
            
    If XSECTION.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[SECTION] = " & (XSECTION.BoundText)
    If xSupp.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xStore.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " [store] = " & MyParn(xStore.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    
    If xage.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If xsex.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
    If XOutlet.Value = 1 Then cWhere = cWhere & Tr(cWhere) & " Outlet = 1 "
    If IsDate(xDate2.Value) Then cWhere = cWhere & Tr(cWhere) & " [DATE] <= " & DateSq(xDate2.Value)
    
    
    If xall(0).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 1 "
    ElseIf xall(1).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 0 "
    End If
    
    If cWhere <> "" Then
        cString = cString & " AND " & cWhere
    End If
    cString = cString & " GROUP BY FILE4_10.CODE, FILE4_10.DESCA, FACT.DESCA, FILE1_10.MOSM, FILE1_10.MODELFACT0, FILE1_10.desca , FILE1_10.MODELNO ORDER BY FILE4_10.CODE, FILE4_10.DESCA, FACT.DESCA, FILE1_10.MOSM, FILE1_10.MODELFACT0 "
    Set data10.Recordset = cmd(cString, con).Execute
End With
FIXGRID0
If grid0.Rows > 0 Then grid0.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub
Sub FIXGRID0()
Dim colRate As Long, ColSales As Long, ColNumbers As Long, ColNet As Long
Dim colCost1 As Long, ColCost2
With grid0
    .RowHeight(0) = 800
    .WordWrap = True
    .ColHidden(1) = True
        
    .TextMatrix(0, 0) = "«·„Ê—œ"
    .ColWidth(0) = 1300
    
    .TextMatrix(0, 1) = "«·„’‰⁄"
    .ColWidth(1) = 0
    
    .TextMatrix(0, 2) = "„Ê”„"
    .ColWidth(2) = 600
    
    .TextMatrix(0, 3) = "„ÊœÌ·"
    .ColWidth(3) = 1500
    
    .TextMatrix(0, 4) = "«·’‰›"
    .ColWidth(4) = 2000
    
    .TextMatrix(0, 5) = "ﬂ„Ì… Ê«—œ"
    .ColWidth(5) = 800
     ColNumbers = 5
    
    Dim col As Integer
    col = 5
    If bCost Then
        col = col + 1
        .TextMatrix(0, col) = "ﬁÌ„… «·Ê«—œ »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1100
        .ColHidden(col) = Not bOpt10
    End If
    
    If bPrice Then
        col = col + 1
        .TextMatrix(0, col) = "«·Ê«—œ »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1100
    End If
    
    col = col + 1
    .TextMatrix(0, col) = "ﬂ„Ì… ’«œ—"
    .ColWidth(col) = 800
    
    If bCost Then
        col = col + 1
        .TextMatrix(0, col) = "’«œ— »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1100
        .ColHidden(col) = Not bOpt10
    End If
    
    If bPrice Then
        col = col + 1
        .TextMatrix(0, col) = "’«œ— »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1100
    End If
    
    col = col + 1
    ColNet = col
    .TextMatrix(0, col) = "’«›Ï «·ﬂ„Ì…"
    .ColWidth(col) = 800
        
    If bCost Then
        col = col + 1
        .TextMatrix(0, col) = "’«›Ï »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1100
        .ColHidden(col) = Not bOpt10
    End If
        
    If bPrice Then
        col = col + 1
        .TextMatrix(0, col) = "’«›Ï »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1100
    End If
        
    col = col + 1
    .TextMatrix(0, col) = "‰”»… „»Ì⁄« "
    .ColWidth(col) = 800
     colRate = col
    
    col = col + 1
    .TextMatrix(0, col) = "⁄œœ „»Ì⁄« "
    .ColWidth(col) = 800
     ColSales = col

    If bCost Then
        col = col + 1
        .TextMatrix(0, col) = " ﬂ·›… „»Ì⁄« "
        .ColWidth(col) = 1100
        .ColHidden(col) = Not bOpt10
    End If
    
    If bPrice Then
        col = col + 1
        .TextMatrix(0, col) = "ﬁÌ„… „»Ì⁄« "
        .ColWidth(col) = 1100
    End If
    
    col = col + 1
    .TextMatrix(0, col) = "—’Ìœ ⁄œœ"
    .ColWidth(col) = 800
    
    If bCost Then
        col = col + 1
        .TextMatrix(0, col) = "—’Ìœ »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1100
        .ColHidden(col) = Not bOpt10
    End If
    
    If bPrice Then
        col = col + 1
        .TextMatrix(0, col) = "—’Ìœ »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1100
        .ColHidden(col) = Not bOpt10
    End If
    
    col = col + 1
    .TextMatrix(0, col) = "«· ﬂ·›…"
    .ColWidth(col) = 1100
    .ColHidden(col) = Not bOpt10
            
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
            
    For Row = 1 To .Rows - 1
        If .ValueMatrix(Row, ColNet) <> 0 Then .TextMatrix(Row, colRate) = Round((.ValueMatrix(Row, ColSales) / .ValueMatrix(Row, ColNet)) * 100, 2)
    Next
            
    .SubtotalPosition = flexSTAbove
    
    For i = ColNumbers To .Cols - 1
        .Subtotal flexSTSum, -1, i, "#0", vbRed, vbYellow, True, "  "
    Next i
    
    For Row = 1 To .Rows - 1
        If .ValueMatrix(Row, ColNet) <> 0 Then .TextMatrix(Row, colRate) = Round((.ValueMatrix(Row, ColSales) / .ValueMatrix(Row, ColNet)) * 100, 2)
        If Row > 1 Then
            If .ValueMatrix(Row, colCost1) <> .ValueMatrix(Row, ColCost2) Then .Cell(flexcpBackColor, Row, colCost1, ColCost2, ColCost2) = vbGreen
        End If
    Next
        
    If .Rows > 1 Then
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End If
End With
End Sub
Private Sub Form_Resize()
SSTab1.Height = Me.Height - SSTab1.Top - 1000
grid1.Height = SSTab1.Height - 500
GRID2.Height = SSTab1.Height - 500
GRID3.Height = SSTab1.Height - 500
GRID4.Height = SSTab1.Height - 500
GRID5.Height = SSTab1.Height - 500
GRID6.Height = SSTab1.Height - 500
GRID7.Height = SSTab1.Height - 500
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
closeCon con
Set VsTInpit_Tab = Nothing
End Sub

Private Sub GRID7_DBLClick()
    VsAllTrans.cGrModel = GRID7.TextMatrix(GRID7.Row, 3)
    VsAllTrans.Show
End Sub

Private Sub optPrice_Click(Index As Integer)
bCost = optPrice(0).Value Or optPrice(1).Value
bPrice = optPrice(0).Value Or optPrice(2).Value
End Sub
Private Sub xModelFact_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    ModelLookupAll Me, oSearchItem
End If
End Sub
Private Sub xGRMODEL_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    GrModelLookupAll Me, oSearchItem
End If
End Sub
Private Sub xSection_LostFocus()
If XSECTION.BoundText = "" Then
    Set DATA3.Recordset = mycmd("Select Code,DescA From File1_50 ORDER BY DESCA", con)
Else
    Set DATA3.Recordset = mycmd("Select Code,DescA From File1_50 where [group] = " & Val(XSECTION.BoundText) & " ORDER BY DESCA", con)
End If
Set xGroup.RowSource = DATA3
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"
End Sub
Private Sub xSupp_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then SuppLookupAll Me, oSearch
End Sub
Sub myProc()
If ActiveControl.Name = xModelFact.Name Then
    xModelFact.text = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 2)
    Unload oSearchItem
End If
If ActiveControl.Name = XGRMODEL.Name Then
    XGRMODEL.text = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    Unload oSearchItem
End If

If ActiveControl.Name = xSupp.Name Then
    xSupp.BoundText = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    oSearch.Hide
End If
If ActiveControl.Name = xFact.Name Then
    xFact.BoundText = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
End If
End Sub
Private Sub xFACT_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From FACT"
Generalarray(2) = "Order by DESCA"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·Ê’›"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·Ê’›"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "≈” ⁄·«„ "
Search3.Show 1

End If
End Sub
Private Sub MYLOAD1()
Dim cString  As String
Dim cWhere As String, cWhereField As String, cField As String
With grid1
    If IsDate(xDate1.Value) Then
        cWhereField = " AND DATE >= " & DateSq(xDate1.Value)
    End If
    
    If chkSection.Value = 1 Then
        cField = cField & "," & "FILE1_10.[SECTION],FILE1_10SC.DESCA"
    End If
    
    If chkGroup.Value = 1 Then
        cField = cField & "," & "FILE1_10.[GROUP],FILE1_50.DESCA"
    End If
    
    If chkMosm.Value = 1 Then
        cField = cField & "," & "FILE1_10.MOSM"
    End If
    
    If chkType.Value = 1 Then
        cField = cField & "," & "FILE1_10.MODELSEX,FILE0_81.DESCA"
    End If
    
    If chkAge.Value = 1 Then
        cField = cField & "," & "FILE1_10.MODELAGE,FILE0_82.DESCA"
    End If
                
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "[IN]")
    
    If bCost Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "FILE1_11_ALL.[IN] * FILE1_11_ALL.COST")
    End If
    
    If bPrice Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "FILE1_11_ALL.[IN] * FILE1_11_ALL.PRICE_SALES")
    End If
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "[OUT] ")
    
    If bCost Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "FILE1_11_ALL.[OUT] * FILE1_11_ALL.COST")
    End If
    
    If bPrice Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "FILE1_11_ALL.[OUT] * FILE1_11_ALL.PRICE_SALES")
    End If
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) ")
    
    If bCost Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) * FILE1_11_ALL.COST")
    End If
    
    If bPrice Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) * FILE1_11_ALL.PRICE_SALES")
    End If
    
    cField = cField & ",0," & _
            myiif("(type = '6')" & cWhereField, "([OUT])")
    
    If bCost Then
        cField = cField & "," & _
                myiif_r("(type = '6')" & cWhereField, "([OUT] * file1_11_all.cost) ")
    End If
    If bPrice Then
        cField = cField & "," & _
                myiif_r("(type = '6')" & cWhereField, "([OUT] * FILE1_11_ALL.PRICE_SALES) ")
    End If
    
    cField = cField & "," & _
            "SUM([IN]-[OUT])"
    
    If bCost Then
        cField = cField & "," & _
                 " ROUND(SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ),2)"
    End If
    
    If bPrice Then
        cField = cField & "," & _
                " ROUND(SUM(([IN]-[OUT]) * FILE1_11_ALL.PRICE_SALES ),2)"
    End If
                                                              
    cString = " SELECT  FILE4_10.CODE , FILE4_10.DESCA,FILE1_10.FACT,FACT.DESCA " & _
              cField & "," & _
              " COUNT(DISTINCT FILE1_10.MODEL) " & _
              " FROM  FILE1_10 INNER JOIN FILE4_10 ON FILE1_10.code = FILE4_10.CODE" & _
              " INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE" & _
              " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
              " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE"
    
    If chkSection.Value = 1 Then
        cString = cString & _
                  " LEFT JOIN FILE1_10SC ON FILE1_10.[SECTION] = FILE1_10SC.CODE"
    End If
    
    If chkGroup.Value = 1 Then
        cString = cString & _
                  " LEFT JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.CODE"
    End If
    
    If chkType.Value = 1 Then
        cString = cString & _
                  " LEFT JOIN FILE0_81 ON FILE0_81.CODE = FILE1_10.MODELSEX"
    End If
              
    If chkAge.Value = 1 Then
        cString = cString & _
                  " LEFT JOIN FILE0_82 ON FILE0_82.CODE = FILE1_10.MODELAGE"
    End If
                                                                          
    cString = cString & " WHERE file1_10.ISNOITEM = 0 "
    
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    If xMosm.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    If xmosm20.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm20.text)
    
    If XSECTION.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[SECTION] = " & (XSECTION.BoundText)
    If xSupp.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xStore.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " [store] = " & MyParn(xStore.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    If XOutlet.Value = 1 Then cWhere = cWhere & Tr(cWhere) & " Outlet = 1 "
    
    If xage.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If xsex.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)

    If IsDate(xDate2.Value) Then cWhere = cWhere & Tr(cWhere) & " [DATE] <= " & DateSq(xDate2.Value)
    
    If xall(0).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 1 "
    ElseIf xall(1).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 0 "
    End If
            
    If cWhere <> "" Then
        cString = cString & " AND " & cWhere
    End If
        
    cString = cString & " GROUP BY FILE4_10.CODE , FILE4_10.DESCA,FILE1_10.FACT,FACT.DESCA"
    
    If chkSection.Value = 1 Then
        cString = cString & "," & "FILE1_10.[SECTION]" & "," & "FILE1_10SC.DESCA"
    End If
    
    If chkGroup.Value = 1 Then
        cString = cString & "," & "FILE1_10.[GROUP],FILE1_50.DESCA"
    End If
    
    If chkMosm.Value = 1 Then
        cString = cString & "," & "FILE1_10.MOSM"
    End If
    
    If chkType.Value = 1 Then
        cString = cString & "," & "FILE1_10.MODELSEX,FILE0_81.DESCA"
    End If
    
    If chkAge.Value = 1 Then
        cString = cString & "," & "FILE1_10.MODELAGE,FILE0_82.DESCA"
    End If
    
    cString = cString & " ORDER BY FILE4_10.CODE,FILE1_10.FACT"
    If chkSection.Value = 1 Then
        cSrting = cString & "," & "FILE1_10.[SECTION]"
    End If
    
    If chkGroup.Value = 1 Then
        cSrting = cString & "," & "FILE1_10.[GROUP]"
    End If
    
    If chkMosm.Value = 1 Then
        cString = cString & "," & "FILE1_10.MOSM"
    End If
    
    If chkType.Value = 1 Then
        cString = cString & "," & "FILE1_10.MODELSEX"
    End If
    
    If chkAge.Value = 1 Then
        cString = cString & "," & "FILE1_10.MODELAGE"
    End If
    
    Set DATA11.Recordset = cmd(cString, con).Execute
End With
FixGrid1
If grid1.Rows > 0 Then grid1.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub
Sub FixGrid1()
Dim col As Long
Dim colRate As Long, ColSales As Long, ColNumbers As Long
With grid1
    .RowHeight(0) = 1000
    .WordWrap = True
    
        
    .TextMatrix(0, col) = "«·„Ê—œ"
    .ColHidden(col) = True
    
    col = col + 1
    .TextMatrix(0, col) = "«·„Ê—œ"
    .ColWidth(col) = 2000
    
    col = col + 1
    .TextMatrix(0, col) = "«·„’‰⁄"
    .ColHidden(col) = True
    
    col = col + 1
    .TextMatrix(0, col) = "«·„’‰⁄"
    .ColWidth(col) = 2000
    
    If chkSection.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·ﬁ”„"
        .ColHidden(col) = True
        
        col = col + 1
        .TextMatrix(0, col) = "«·ﬁ”„"
        .ColWidth(col) = 2000
    End If
    
    If chkGroup.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·„Ã„Ê⁄…"
        .ColHidden(col) = True
        
        col = col + 1
        .TextMatrix(0, col) = "«·„Ã„Ê⁄…"
        .ColWidth(col) = 2000
    End If
    
    If chkSection.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·ﬁ”„"
        .ColHidden(col) = True
        
        col = col + 1
        .TextMatrix(0, col) = "«·ﬁ”„"
        .ColWidth(col) = 2000
    End If
    
    If chkMosm.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·„Ê”„"
        .ColWidth(col) = 1000
    End If
    
    If chkType.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·‰Ê⁄"
        .ColHidden(col) = True
            
        col = col + 1
        .TextMatrix(0, col) = "«·‰Ê⁄"
        .ColWidth(col) = 1000
    End If
    
    If chkAge.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·”‰"
        .ColHidden(col) = True
        
        col = col + 1
        .TextMatrix(0, col) = "«·”‰"
        .ColWidth(col) = 1200
    End If
     
     col = col + 1
     ColNumbers = col
    .TextMatrix(0, col) = "ﬂ„Ì… Ê«—œ"
    .ColWidth(col) = 1000
    .ColDataType(col) = flexDTDouble
    
    If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "Ê«—œ »”⁄— «· ﬂ·›…"
        .ColHidden(col) = Not bOpt10
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
    If bPrice Then
         col = col + 1
        .TextMatrix(0, col) = "Ê«—œ »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
     col = col + 1
    .TextMatrix(0, col) = "ﬂ„Ì… ’«œ—"
    .ColWidth(col) = 1000
    .ColDataType(col) = flexDTDouble

    If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "’«œ— »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
        .ColDataType(col) = flexDTDouble
    End If
    
    If bPrice Then
         col = col + 1
        .TextMatrix(0, col) = "’«œ— »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
     col = col + 1
     ColNet = col
    .TextMatrix(0, col) = "’«›Ï «·ﬂ„Ì…"
    .ColWidth(col) = 1000
    .ColDataType(col) = flexDTDouble
    
    If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "’«›Ï «·ﬂ„Ì… »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
        .ColDataType(col) = flexDTDouble
    End If
    
    If bPrice Then
         col = col + 1
        .TextMatrix(0, col) = "’«›Ï «·ﬂ„Ì… »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
     col = col + 1
     colRate = col
    .TextMatrix(0, col) = "‰”»… „»Ì⁄« "
    .ColWidth(col) = 1300
    .ColDataType(col) = flexDTDouble
    
     col = col + 1
     ColSales = col
    .TextMatrix(0, col) = "⁄œœ „»Ì⁄« "
    .ColWidth(col) = 1000
    .ColDataType(col) = flexDTDouble
     
     If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "„»Ì⁄«  »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
        .ColDataType(col) = flexDTDouble
    End If
     
    If bPrice Then
         col = col + 1
        .TextMatrix(0, col) = "„»Ì⁄«  »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
     col = col + 1
    .TextMatrix(0, col) = "—’Ìœ ⁄œœ"
    .ColWidth(col) = 1300
    .ColDataType(col) = flexDTDouble

    
    If bCost Then
        col = col + 1
        .TextMatrix(0, col) = "ﬁÌ„… «·—’Ìœ »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
        .ColDataType(col) = flexDTDouble
    End If
    
    If bPrice Then
        col = col + 1
        .TextMatrix(0, col) = "ﬁÌ„… «·—’Ìœ »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
    col = col + 1
    .TextMatrix(0, col) = "⁄œœ „ÊœÌ·« "
    .ColWidth(col) = 1300
    .ColDataType(col) = flexDTDouble

            
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    For Row = 1 To .Rows - 1
        If .ValueMatrix(Row, ColNet) <> 0 Then .TextMatrix(Row, colRate) = Round((.ValueMatrix(Row, ColSales) / .ValueMatrix(Row, ColNet)) * 100, 2)
    Next
    
    .SubtotalPosition = flexSTAbove
    For i = ColNumbers To .Cols - 1
        .Subtotal flexSTSum, -1, i, "#0", vbRed, vbYellow, True, "  "
    Next
    
    If .Rows > 1 Then
        If Val(.TextMatrix(1, ColNet)) <> 0 Then .TextMatrix(1, colRate) = Round((Val(.TextMatrix(1, ColSales)) / Val(.TextMatrix(1, ColNet))) * 100, 2)
    End If
    
    If .Rows > 1 Then
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End If
End With
End Sub
Private Sub MYLOAD2()
Dim cString  As String
Dim cWhere As String, cWhereField As String, cField As String
With GRID2
    
    If IsDate(xDate1.Value) Then
        cWhereField = " AND DATE >= " & DateSq(xDate1.Value)
    End If
                
    If chkGroup.Value = 1 Then
        cField = cField & "," & "FILE1_10.[GROUP],FILE1_50.DESCA"
    End If
             
    If chkMosm.Value = 1 Then
        cField = cField & "," & "FILE1_10.MOSM"
    End If
        
    If chkType.Value = 1 Then
        cField = cField & "," & "FILE1_10.MODELSEX,FILE0_81.DESCA"
    End If
    
    If chkAge.Value = 1 Then
        cField = cField & "," & "FILE1_10.MODELAGE,FILE0_82.DESCA"
    End If
        
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "[IN]")
    
    If bCost Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "FILE1_11_ALL.[IN] * FILE1_11_ALL.COST")
    End If
    
    If bPrice Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "FILE1_11_ALL.[IN] * FILE1_11_ALL.PRICE_SALES")
    End If
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "[OUT] ")
    
    If bCost Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "FILE1_11_ALL.[OUT] * FILE1_11_ALL.COST")
    End If
    
    If bPrice Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "FILE1_11_ALL.[OUT] * FILE1_11_ALL.PRICE_SALES")
    End If
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) ")
    
    If bCost Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) * FILE1_11_ALL.COST")
    End If
    
    If bPrice Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) * FILE1_11_ALL.PRICE_SALES  ")
    End If
    
    cField = cField & ",0," & _
            myiif("(type = '6')" & cWhereField, "([OUT])")
    
    If bCost Then
        cField = cField & "," & _
                myiif_r("(type = '6')" & cWhereField, "([OUT] * file1_11_all.cost) ")
    End If
                
    If bPrice Then
        cField = cField & "," & _
                myiif_r("(type = '6')" & cWhereField, "([OUT] * FILE1_11_ALL.PRICE_SALES) ")
    End If
    
    cField = cField & "," & _
            "SUM([IN]-[OUT])"
    
    If bCost Then
        cField = cField & "," & _
                 " ROUND(SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ),2)"
    End If
    
    If bPrice Then
        cField = cField & "," & _
                " ROUND(SUM(([IN]-[OUT]) * FILE1_11_ALL.PRICE_SALES ),2)"
    End If
        
    cString = " SELECT  FILE1_10.SECTION, FILE1_10SC.DESCA " & _
              cField & "," & _
              " COUNT(DISTINCT FILE1_10.MODEL) " & _
              " FROM  FILE1_10 INNER JOIN FILE1_10SC ON FILE1_10.[SECTION] = FILE1_10SC.CODE" & _
              " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
              " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE"
        
    If chkGroup.Value = 1 Then
        cString = cString & _
                  " LEFT JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.CODE"
    End If
        
    If chkType.Value = 1 Then
        cString = cString & _
                  " LEFT JOIN FILE0_81 ON FILE0_81.CODE = FILE1_10.MODELSEX"
    End If
              
    If chkAge.Value = 1 Then
        cString = cString & _
                  " LEFT JOIN FILE0_82 ON FILE0_82.CODE = FILE1_10.MODELAGE"
    End If
                                                                          
    cString = cString & " WHERE file1_10.ISNOITEM = 0 "
    
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    If xMosm.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    If xmosm20.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm20.text)
    
    If XSECTION.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[SECTION] = " & (XSECTION.BoundText)
    If xSupp.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xStore.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " [store] = " & MyParn(xStore.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    If XOutlet.Value = 1 Then cWhere = cWhere & Tr(cWhere) & " Outlet = 1 "
    
    If xage.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If xsex.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)

    If IsDate(xDate2.Value) Then cWhere = cWhere & Tr(cWhere) & " [DATE] <= " & DateSq(xDate2.Value)
    
    If xall(0).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 1 "
    ElseIf xall(1).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 0 "
    End If
            
    If cWhere <> "" Then
        cString = cString & " AND " & cWhere
    End If
        
    cString = cString & " GROUP BY FILE1_10.[SECTION], FILE1_10SC.DESCA"
        
        
    If chkGroup.Value = 1 Then
        cString = cString & "," & "FILE1_10.[GROUP],FILE1_50.DESCA"
    End If
        
    If chkMosm.Value = 1 Then
        cString = cString & "," & "FILE1_10.MOSM"
    End If
    
    If chkType.Value = 1 Then
        cString = cString & "," & "FILE1_10.MODELSEX,FILE0_81.DESCA"
    End If
    
    If chkAge.Value = 1 Then
        cString = cString & "," & "FILE1_10.MODELAGE,FILE0_82.DESCA"
    End If
    
    cString = cString & " ORDER BY FILE1_10.[SECTION]"
    
    If chkGroup.Value = 1 Then
        cString = cString & "," & "FILE1_10.[GROUP]"
    End If
    
    If chkMosm.Value = 1 Then
        cString = cString & "," & "FILE1_10.MOSM"
    End If
    
    If chkType.Value = 1 Then
        cString = cString & "," & "FILE1_10.MODELSEX"
    End If
    
    If chkAge.Value = 1 Then
        cString = cString & "," & "FILE1_10.MODELAGE"
    End If
    
    Set data12.Recordset = cmd(cString, con).Execute
End With
FixGrid2
If GRID2.Rows > 0 Then GRID2.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub
Sub FixGrid2()
Dim col As Long
Dim colRate As Long, ColSales As Long, ColNumbers As Long

With GRID2
    .RowHeight(0) = 1000
    .WordWrap = True
    
        
    .TextMatrix(0, col) = "«·ﬁ”„"
    .ColHidden(col) = True
    
    col = col + 1
    .TextMatrix(0, col) = "«·ﬁ”„"
    .ColWidth(col) = 2000
            
    If chkGroup.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·„Ã„Ê⁄…"
        .ColHidden(col) = True
        
        col = col + 1
        .TextMatrix(0, col) = "«·„Ã„Ê⁄…"
        .ColWidth(col) = 2000
    End If
            
    If chkMosm.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·„Ê”„"
        .ColWidth(col) = 1000
    End If
    
    If chkType.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·‰Ê⁄"
        .ColHidden(col) = True
            
        col = col + 1
        .TextMatrix(0, col) = "«·‰Ê⁄"
        .ColWidth(col) = 1000
    End If
    
    If chkAge.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·”‰"
        .ColHidden(col) = True
        
        col = col + 1
        .TextMatrix(0, col) = "«·”‰"
        .ColWidth(col) = 1200
    End If
     
     col = col + 1
     ColNumbers = col
    .TextMatrix(0, col) = "ﬂ„Ì… Ê«—œ"
    .ColWidth(col) = 1000
    .ColDataType(col) = flexDTDouble
    
    If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "Ê«—œ »”⁄— «· ﬂ·›…"
        .ColHidden(col) = Not bOpt10
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
    If bPrice Then
         col = col + 1
        .TextMatrix(0, col) = "Ê«—œ »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
     col = col + 1
    .TextMatrix(0, col) = "ﬂ„Ì… ’«œ—"
    .ColWidth(col) = 1000
    .ColDataType(col) = flexDTDouble

    If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "’«œ— »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
        .ColDataType(col) = flexDTDouble
    End If
     
    If bPrice Then
         col = col + 1
        .TextMatrix(0, col) = "’«œ— »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
     col = col + 1
     ColNet = col
    .TextMatrix(0, col) = "’«›Ï «·ﬂ„Ì…"
    .ColWidth(col) = 1000
    .ColDataType(col) = flexDTDouble
    
    If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "’«›Ï «·ﬂ„Ì… »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
        .ColDataType(col) = flexDTDouble
    End If
            
    If bPrice Then
         col = col + 1
        .TextMatrix(0, col) = "’«›Ï «·ﬂ„Ì… »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
     
     col = col + 1
     colRate = col
    .TextMatrix(0, col) = "‰”»… „»Ì⁄« "
    .ColWidth(col) = 1300
    .ColDataType(col) = flexDTDouble

     col = col + 1
     ColSales = col
    .TextMatrix(0, col) = "⁄œœ „»Ì⁄« "
    .ColWidth(col) = 1000
    .ColDataType(col) = flexDTDouble
     
    If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "„»Ì⁄«  »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
        .ColDataType(col) = flexDTDouble
    End If
    
    If bPrice Then
         col = col + 1
        .TextMatrix(0, col) = "„»Ì⁄«  »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
     col = col + 1
    .TextMatrix(0, col) = "—’Ìœ ⁄œœ"
    .ColWidth(col) = 1300
    .ColDataType(col) = flexDTDouble

    If bCost Then
        col = col + 1
        .TextMatrix(0, col) = "ﬁÌ„… «·—’Ìœ »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
        .ColDataType(col) = flexDTDouble
    End If
    
    If bPrice Then
        col = col + 1
        .TextMatrix(0, col) = "ﬁÌ„… «·—’Ìœ »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
    col = col + 1
    .TextMatrix(0, col) = "⁄œœ „ÊœÌ·« "
    .ColWidth(col) = 1300
    .ColDataType(col) = flexDTDouble

            
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    
    For Row = 1 To .Rows - 1
        If .ValueMatrix(Row, ColNet) <> 0 Then .TextMatrix(Row, colRate) = Round((.ValueMatrix(Row, ColSales) / .ValueMatrix(Row, ColNet)) * 100, 2)
    Next
    
    .SubtotalPosition = flexSTAbove
    For i = ColNumbers To .Cols - 1
        .Subtotal flexSTSum, -1, i, "#0", vbRed, vbYellow, True, "  "
    Next
    
    If .Rows > 1 Then
        If Val(.TextMatrix(1, ColNet)) <> 0 Then .TextMatrix(1, colRate) = Round((Val(.TextMatrix(1, ColSales)) / Val(.TextMatrix(1, ColNet))) * 100, 2)
    End If
    
    If .Rows > 1 Then
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End If
    End With
End Sub
Private Sub MYLOAD3()
Dim cString  As String
Dim cWhere As String, cWhereField As String, cField As String
With GRID3
    If IsDate(xDate1.Value) Then
        cWhereField = " AND DATE >= " & DateSq(xDate1.Value)
    End If
        
    If chkSection.Value = 1 Then
        cField = cField & "," & "FILE1_10.[SECTION]" & "," & "FILE1_10SC.DESCA"
    End If
                
    If chkMosm.Value = 1 Then
        cField = cField & "," & "FILE1_10.MOSM"
    End If
    
    If chkType.Value = 1 Then
        cField = cField & "," & "FILE1_10.MODELSEX,FILE0_81.DESCA"
    End If
    
    If chkAge.Value = 1 Then
        cField = cField & "," & "FILE1_10.MODELAGE,FILE0_82.DESCA"
    End If
        
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "[IN]")
    
    If bCost Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "FILE1_11_ALL.[IN] * FILE1_11_ALL.COST")
    End If
    
    If bPrice Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "FILE1_11_ALL.[IN] * FILE1_11_ALL.PRICE_SALES")
    End If
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "[OUT] ")
    
    If bCost Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "FILE1_11_ALL.[OUT] * FILE1_11_ALL.COST")
    End If
    
    If bPrice Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "FILE1_11_ALL.[OUT] * FILE1_11_ALL.PRICE_SALES")
    End If
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) ")
    
    If bCost Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) * FILE1_11_ALL.COST")
    End If
    
    If bPrice Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) * FILE1_11_ALL.PRICE_SALES")
    End If
    
    cField = cField & ",0," & _
            myiif("(type = '6')" & cWhereField, "([OUT])")
    
    If bCost Then
        cField = cField & "," & _
                myiif_r("(type = '6')" & cWhereField, "([OUT] * file1_11_all.cost) ")
    End If
    
    If bPrice Then
        cField = cField & "," & _
                myiif_r("(type = '6')" & cWhereField, "([OUT] * FILE1_11_ALL.PRICE_SALES) ")
    End If
    
    cField = cField & "," & _
            "SUM([IN]-[OUT])"
    
    If bCost Then
        cField = cField & "," & _
                 " ROUND(SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ),2)"
    End If
    
    If bPrice Then
        cField = cField & "," & _
                " ROUND(SUM(([IN]-[OUT]) * FILE1_11_ALL.PRICE_SALES ),2)"
    End If
    
    cString = " SELECT  FILE1_10.[GROUP], FILE1_50.DESCA " & _
              cField & _
              ", COUNT(DISTINCT FILE1_10.MODEL) " & _
              " FROM  FILE1_10 INNER JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.[CODE]" & _
              " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
              " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE"
    
    If chkSection.Value = 1 Then
        cString = cString & _
                  " LEFT JOIN FILE1_10SC ON FILE1_10.[SECTION] = FILE1_10SC.CODE"
    End If
        
    If chkType.Value = 1 Then
        cString = cString & _
                  " LEFT JOIN FILE0_81 ON FILE0_81.CODE = FILE1_10.MODELSEX"
    End If
              
    If chkAge.Value = 1 Then
        cString = cString & _
                  " LEFT JOIN FILE0_82 ON FILE0_82.CODE = FILE1_10.MODELAGE"
    End If
                                                                          
    cString = cString & " WHERE file1_10.ISNOITEM = 0 "
    
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    If xMosm.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    If xmosm20.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm20.text)
    
    If XSECTION.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[SECTION] = " & (XSECTION.BoundText)
    If xSupp.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xStore.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " [store] = " & MyParn(xStore.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    If XOutlet.Value = 1 Then cWhere = cWhere & Tr(cWhere) & " Outlet = 1 "
    
    If xage.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If xsex.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)

    If IsDate(xDate2.Value) Then cWhere = cWhere & Tr(cWhere) & " [DATE] <= " & DateSq(xDate2.Value)
    
    If xall(0).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 1 "
    ElseIf xall(1).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 0 "
    End If
            
    If cWhere <> "" Then
        cString = cString & " AND " & cWhere
    End If
        
    cString = cString & " GROUP BY FILE1_10.[GROUP], FILE1_50.DESCA"
        
    If chkSection.Value = 1 Then
        cString = cString & "," & "FILE1_10.[SECTION]" & "," & "FILE1_10SC.DESCA"
    End If
        
    If chkMosm.Value = 1 Then
        cString = cString & "," & "FILE1_10.MOSM"
    End If
    
    If chkType.Value = 1 Then
        cString = cString & "," & "FILE1_10.MODELSEX,FILE0_81.DESCA"
    End If
    
    If chkAge.Value = 1 Then
        cString = cString & "," & "FILE1_10.MODELAGE,FILE0_82.DESCA"
    End If
    
    cString = cString & " ORDER BY FILE1_10.[GROUP]"
    
    If chkSection.Value = 1 Then
        cString = cString & "," & "FILE1_10.[SECTION]"
    End If
    
    If chkMosm.Value = 1 Then
        cString = cString & "," & "FILE1_10.MOSM"
    End If
    
    If chkType.Value = 1 Then
        cString = cString & "," & "FILE1_10.MODELSEX"
    End If
    
    If chkAge.Value = 1 Then
        cString = cString & "," & "FILE1_10.MODELAGE"
    End If
    
    Set DATA13.Recordset = cmd(cString, con).Execute
End With
FixGrid3
If GRID3.Rows > 0 Then GRID3.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub
Sub FixGrid3()
Dim col As Long
Dim colRate As Long, ColSales As Long, ColNumbers As Long

With GRID3
    .RowHeight(0) = 1000
    .WordWrap = True
            
    .TextMatrix(0, col) = "«·„Ã„Ê⁄…"
    .ColHidden(col) = True
    
    col = col + 1
    .TextMatrix(0, col) = "«·„Ã„Ê⁄…"
    .ColWidth(col) = 2000
            
    If chkSection.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·ﬁ”„"
        .ColHidden(col) = True
        
        col = col + 1
        .TextMatrix(0, col) = "«·ﬁ”„"
        .ColWidth(col) = 2000
    End If
            
    If chkMosm.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·„Ê”„"
        .ColWidth(col) = 1000
    End If
    
    If chkType.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·‰Ê⁄"
        .ColHidden(col) = True
            
        col = col + 1
        .TextMatrix(0, col) = "«·‰Ê⁄"
        .ColWidth(col) = 1000
    End If
    
    If chkAge.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·”‰"
        .ColHidden(col) = True
        
        col = col + 1
        .TextMatrix(0, col) = "«·”‰"
        .ColWidth(col) = 1200
    End If
     
     col = col + 1
     ColNumbers = col
    .TextMatrix(0, col) = "ﬂ„Ì… Ê«—œ"
    .ColWidth(col) = 1000
    .ColDataType(col) = flexDTDouble
    
    If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "Ê«—œ »”⁄— «· ﬂ·›…"
        .ColHidden(col) = Not bOpt10
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
    If bPrice Then
         col = col + 1
        .TextMatrix(0, col) = "Ê«—œ »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
     
     col = col + 1
    .TextMatrix(0, col) = "ﬂ„Ì… ’«œ—"
    .ColWidth(col) = 1000
    .ColDataType(col) = flexDTDouble
    
    If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "’«œ— »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
        .ColDataType(col) = flexDTDouble
    End If
    
    If bPrice Then
         col = col + 1
        .TextMatrix(0, col) = "’«œ— »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
     col = col + 1
     ColNet = col
    .TextMatrix(0, col) = "’«›Ï «·ﬂ„Ì…"
    .ColWidth(col) = 1000
    .ColDataType(col) = flexDTDouble
         
    If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "’«›Ï «·ﬂ„Ì… »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
        .ColDataType(col) = flexDTDouble
    End If
    
    If bPrice Then
         col = col + 1
        .TextMatrix(0, col) = "’«›Ï «·ﬂ„Ì… »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
          
     col = col + 1
     colRate = col
    .TextMatrix(0, col) = "‰”»… „»Ì⁄« "
    .ColWidth(col) = 1300
    .ColDataType(col) = flexDTDouble
    
     col = col + 1
     ColSales = col
    .TextMatrix(0, col) = "⁄œœ „»Ì⁄« "
    .ColWidth(col) = 1000
    .ColDataType(col) = flexDTDouble
     
    If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "„»Ì⁄«  »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
        .ColDataType(col) = flexDTDouble
    End If
    
    If bPrice Then
         col = col + 1
        .TextMatrix(0, col) = "„»Ì⁄«  »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
     col = col + 1
    .TextMatrix(0, col) = "—’Ìœ ⁄œœ"
    .ColWidth(col) = 1300
    .ColDataType(col) = flexDTDouble
        
    If bCost Then
        col = col + 1
        .TextMatrix(0, col) = "ﬁÌ„… «·—’Ìœ »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
        .ColDataType(col) = flexDTDouble
    End If
    
    If bPrice Then
        col = col + 1
        .TextMatrix(0, col) = "ﬁÌ„… «·—’Ìœ »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
    col = col + 1
    .TextMatrix(0, col) = "⁄œœ „ÊœÌ·« "
    .ColWidth(col) = 1300
    .ColDataType(col) = flexDTDouble

            
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
        
    For Row = 1 To .Rows - 1
        If .ValueMatrix(Row, ColNet) <> 0 Then .TextMatrix(Row, colRate) = Round((.ValueMatrix(Row, ColSales) / .ValueMatrix(Row, ColNet)) * 100, 2)
    Next
    
    .SubtotalPosition = flexSTAbove
    For i = ColNumbers To .Cols - 1
        .Subtotal flexSTSum, -1, i, "#0", vbRed, vbYellow, True, "  "
    Next
    
    If .Rows > 1 Then
        If Val(.TextMatrix(1, ColNet)) <> 0 Then .TextMatrix(1, colRate) = Round((Val(.TextMatrix(1, ColSales)) / Val(.TextMatrix(1, ColNet))) * 100, 2)
    End If
    
    If .Rows > 1 Then
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End If
End With
End Sub
Private Sub MYLOAD4()
Dim cString  As String
Dim cWhere As String, cWhereField As String, cField As String
With GRID4
    If IsDate(xDate1.Value) Then
        cWhereField = " AND DATE >= " & DateSq(xDate1.Value)
    End If
        
    If chkSection.Value = 1 Then
        cField = cField & "," & "FILE1_10.[SECTION],FILE1_10SC.DESCA"
    End If
        
    If chkGroup.Value = 1 Then
        cField = cField & "," & "FILE1_10.[GROUP],FILE1_50.DESCA"
    End If
        
    If chkMosm.Value = 1 Then
        cField = cField & "," & "FILE1_10.MOSM"
    End If
    
    If chkType.Value = 1 Then
        cField = cField & "," & "FILE1_10.MODELSEX,FILE0_81.DESCA"
    End If
            
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "[IN]")
    
    If bCost Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "FILE1_11_ALL.[IN] * FILE1_11_ALL.COST")
    End If
        
    If bPrice Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "FILE1_11_ALL.[IN] * FILE1_11_ALL.PRICE_SALES")
    End If
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "[OUT] ")
    
    If bCost Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "FILE1_11_ALL.[OUT] * FILE1_11_ALL.COST")
    End If
    
    If bPrice Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "FILE1_11_ALL.[OUT] * FILE1_11_ALL.PRICE_SALES")
    End If
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) ")
    
    If bCost Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) * FILE1_11_ALL.COST")
    End If
    
    If bPrice Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) * FILE1_11_ALL.PRICE_SALES  ")
    End If
    
    cField = cField & ",0," & _
            myiif("(type = '6')" & cWhereField, "([OUT])")
    
    If bCost Then
        cField = cField & "," & _
                myiif_r("(type = '6')" & cWhereField, "([OUT] * file1_11_all.cost) ")
    End If
    
    If bPrice Then
        cField = cField & "," & _
                myiif_r("(type = '6')" & cWhereField, "([OUT] * FILE1_11_ALL.PRICE_SALES) ")
    End If
    
    cField = cField & "," & _
            "SUM([IN]-[OUT])"
    
    If bCost Then
        cField = cField & "," & _
                 " ROUND(SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ),2)"
    End If
    
    If bPrice Then
        cField = cField & "," & _
                " ROUND(SUM(([IN]-[OUT]) * FILE1_11_ALL.PRICE_SALES ),2)"
    End If
    
    cString = " SELECT  FILE1_10.[MODELAGE], FILE0_82.DESCA " & _
              cField & _
              " ,COUNT(DISTINCT FILE1_10.MODEL) " & _
              " FROM  FILE1_10 INNER JOIN FILE0_82 ON FILE1_10.[MODELAGE] = FILE0_82.[CODE]" & _
              " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
              " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE"
    
    If chkSection.Value = 1 Then
        cString = cString & _
                  " LEFT JOIN FILE1_10SC ON FILE1_10.[SECTION] = FILE1_10SC.CODE"
    End If
        
    If chkGroup.Value = 1 Then
        cString = cString & _
                  " LEFT JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.CODE"
    End If
        
    If chkType.Value = 1 Then
        cString = cString & _
                  " LEFT JOIN FILE0_81 ON FILE0_81.CODE = FILE1_10.MODELSEX"
    End If
              
'    If chkAge.Value = 1 Then
'        cString = cString & _
'                  " LEFT JOIN FILE0_82 ON FILE0_82.CODE = FILE1_10.MODELAGE"
'    End If
                                                                          
    cString = cString & " WHERE file1_10.ISNOITEM = 0 "
    
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    If xMosm.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    If xmosm20.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm20.text)
    
    If XSECTION.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[SECTION] = " & (XSECTION.BoundText)
    If xSupp.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xStore.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " [store] = " & MyParn(xStore.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    If XOutlet.Value = 1 Then cWhere = cWhere & Tr(cWhere) & " Outlet = 1 "
    
    If xage.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If xsex.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)

    If IsDate(xDate2.Value) Then cWhere = cWhere & Tr(cWhere) & " [DATE] <= " & DateSq(xDate2.Value)
    
    If xall(0).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 1 "
    ElseIf xall(1).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 0 "
    End If
            
    If cWhere <> "" Then
        cString = cString & " AND " & cWhere
    End If
        
    cString = cString & " GROUP BY FILE1_10.[MODELAGE], FILE0_82.DESCA"
        
    If chkSection.Value = 1 Then
        cString = cString & "," & "FILE1_10.[SECTION],FILE1_10SC.DESCA"
    End If
        
    If chkGroup.Value = 1 Then
        cString = cString & "," & "FILE1_10.[GROUP],FILE1_50.DESCA"
    End If
        
    If chkMosm.Value = 1 Then
        cString = cString & "," & "FILE1_10.MOSM"
    End If
    
    If chkType.Value = 1 Then
        cString = cString & "," & "FILE1_10.MODELSEX,FILE0_81.DESCA"
    End If
    
'    If chkAge.Value = 1 Then
'        cString = cString & "," & "FILE1_10.MODELAGE,FILE0_82.DESCA"
'    End If
    
    cString = cString & " ORDER BY FILE1_10.[MODELAGE]"
    
    If chkSection.Value = 1 Then
        cString = cString & "," & "FILE1_10.[SECTION]"
    End If
    
    If chkGroup.Value = 1 Then
        cString = cString & ",FILE1_10.[GROUP]"
    End If
    
    If chkMosm.Value = 1 Then
        cString = cString & ",FILE1_10.MOSM"
    End If
    
    If chkType.Value = 1 Then
        cString = cString & ",FILE1_10.MODELSEX"
    End If
    
'    If chkAge.Value = 1 Then
'        cString = cString & "," & "FILE1_10.MODELAGE"
'    End If
    
    Set DATA14.Recordset = cmd(cString, con).Execute
End With
FIXGRID4
If GRID4.Rows > 0 Then GRID4.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub
Sub FIXGRID4()
Dim col As Long
Dim colRate As Long, ColSales As Long, ColNumbers As Long

With GRID4
    .RowHeight(0) = 1000
    .WordWrap = True
            
    .TextMatrix(0, col) = "«·›∆… «·⁄„—Ì…"
    .ColHidden(col) = True
    
    col = col + 1
    .TextMatrix(0, col) = "«·›∆… «·⁄„—Ì…"
    .ColWidth(col) = 2000
            
    If chkSection.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·ﬁ”„"
        .ColHidden(col) = True
        
        col = col + 1
        .TextMatrix(0, col) = "«·ﬁ”„"
        .ColWidth(col) = 2000
    End If
            
    If chkGroup.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·„Ã„Ê⁄…"
        .ColHidden(col) = True
        
        col = col + 1
        .TextMatrix(0, col) = "«·„Ã„Ê⁄…"
        .ColWidth(col) = 2000
    End If
            
    If chkMosm.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·„Ê”„"
        .ColWidth(col) = 1000
    End If
    
    If chkType.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·‰Ê⁄"
        .ColHidden(col) = True
            
        col = col + 1
        .TextMatrix(0, col) = "«·‰Ê⁄"
        .ColWidth(col) = 1000
    End If
    
     
     col = col + 1
     ColNumbers = col
    .TextMatrix(0, col) = "ﬂ„Ì… Ê«—œ"
    .ColWidth(col) = 1000
    .ColDataType(col) = flexDTDouble
    
    If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "Ê«—œ »”⁄— «· ﬂ·›…"
        .ColHidden(col) = Not bOpt10
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
    If bPrice Then
         col = col + 1
        .TextMatrix(0, col) = "Ê«—œ »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
     col = col + 1
    .TextMatrix(0, col) = "ﬂ„Ì… ’«œ—"
    .ColWidth(col) = 1000
    .ColDataType(col) = flexDTDouble
    
    If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "’«œ— »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
        .ColDataType(col) = flexDTDouble
    End If
    
    If bPrice Then
         col = col + 1
        .TextMatrix(0, col) = "’«œ— »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
     col = col + 1
     ColNet = col
    .TextMatrix(0, col) = "’«›Ï «·ﬂ„Ì…"
    .ColWidth(col) = 1000
    .ColDataType(col) = flexDTDouble
    
    If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "’«›Ï «·ﬂ„Ì… »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
        .ColDataType(col) = flexDTDouble
    End If
    
    If bPrice Then
         col = col + 1
        .TextMatrix(0, col) = "’«›Ï «·ﬂ„Ì… »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
     col = col + 1
     colRate = col
    .TextMatrix(0, col) = "‰”»… „»Ì⁄« "
    .ColWidth(col) = 1300
    .ColDataType(col) = flexDTDouble
    
     col = col + 1
     ColSales = col
    .TextMatrix(0, col) = "⁄œœ „»Ì⁄« "
    .ColWidth(col) = 1000
    .ColDataType(col) = flexDTDouble
     
    If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "„»Ì⁄«  »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
        .ColDataType(col) = flexDTDouble
    End If
    
     
    If bPrice Then
         col = col + 1
        .TextMatrix(0, col) = "„»Ì⁄«  »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
     col = col + 1
    .TextMatrix(0, col) = "—’Ìœ ⁄œœ"
    .ColWidth(col) = 1300
    .ColDataType(col) = flexDTDouble

    If bCost Then
        col = col + 1
        .TextMatrix(0, col) = "ﬁÌ„… «·—’Ìœ »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
        .ColDataType(col) = flexDTDouble
    End If
    
    If bPrice Then
        col = col + 1
        .TextMatrix(0, col) = "ﬁÌ„… «·—’Ìœ »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
    col = col + 1
    .TextMatrix(0, col) = "⁄œœ „ÊœÌ·« "
    .ColWidth(col) = 1300
    .ColDataType(col) = flexDTDouble

            
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
     
    For Row = 1 To .Rows - 1
        If .ValueMatrix(Row, ColNet) <> 0 Then .TextMatrix(Row, colRate) = Round((.ValueMatrix(Row, ColSales) / .ValueMatrix(Row, ColNet)) * 100, 2)
    Next
    
    .SubtotalPosition = flexSTAbove
    For i = ColNumbers To .Cols - 1
        .Subtotal flexSTSum, -1, i, "#0", vbRed, vbYellow, True, "  "
    Next
    
    If .Rows > 1 Then
        If Val(.TextMatrix(1, ColNet)) <> 0 Then .TextMatrix(1, colRate) = Round((Val(.TextMatrix(1, ColSales)) / Val(.TextMatrix(1, ColNet))) * 100, 2)
    End If
    
    If .Rows > 1 Then
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End If
End With
End Sub
Private Sub myload6()
Dim cString  As String
Dim cWhere As String, cWhereField As String, cField As String
With GRID6
    If IsDate(xDate1.Value) Then
        cWhereField = " AND DATE >= " & DateSq(xDate1.Value)
    End If
                
    If chkSection.Value = 1 Then
        cField = cField & "," & "FILE1_10.[SECTION],FILE1_10SC.DESCA"
    End If
    
    If chkGroup.Value = 1 Then
        cField = cField & "," & "FILE1_10.[GROUP],FILE1_50.DESCA"
    End If
    
        
    If chkMosm.Value = 1 Then
        cField = cField & "," & "FILE1_10.MOSM"
    End If
    
    If chkType.Value = 1 Then
        cField = cField & "," & "FILE1_10.MODELSEX,FILE0_81.DESCA"
    End If
    
    If chkAge.Value = 1 Then
        cField = cField & "," & "FILE1_10.MODELAGE,FILE0_82.DESCA"
    End If
        
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "[IN]")
    
    If bCost Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "FILE1_11_ALL.[IN] * FILE1_11_ALL.COST")
    End If
        
    cField = cField & "," & _
            myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "FILE1_11_ALL.[IN] * FILE1_11_ALL.PRICE_SALES")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "[OUT] ")
    
    If bCost Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "FILE1_11_ALL.[OUT] * FILE1_11_ALL.COST")
    End If
    
    If bPrice Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "FILE1_11_ALL.[OUT] * FILE1_11_ALL.PRICE_SALES")
    End If
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) ")
    
    If bCost Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) * FILE1_11_ALL.COST")
    End If
    
    If bPrice Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) * FILE1_11_ALL.PRICE_SALES")
    End If
    
    cField = cField & ",0," & _
            myiif("(type = '6')" & cWhereField, "([OUT])")
    
    If bCost Then
        cField = cField & "," & _
                myiif_r("(type = '6')" & cWhereField, "([OUT] * file1_11_all.cost) ")
                
    End If
    
    If bPrice Then
        cField = cField & "," & _
                myiif_r("(type = '6')" & cWhereField, "([OUT] * FILE1_11_ALL.PRICE_SALES) ")
    End If
    
    cField = cField & "," & _
            "SUM([IN]-[OUT])"
    
    If bCost Then
        cField = cField & "," & _
                 " ROUND(SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ),2)"
    End If
    
    If bPrice Then
        cField = cField & "," & _
                " ROUND(SUM(([IN]-[OUT]) * FILE1_11_ALL.PRICE_SALES ),2)"
    End If
    
    cString = " SELECT  FILE1_11_ALL.STORE , STORE_ALL.DESCA " & _
              cField & _
              ",COUNT(DISTINCT FILE1_10.MODEL) AS COUNT " & _
              " FROM  FILE1_10 INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
              " INNER JOIN STORE_ALL ON STORE_ALL.CODE = FILE1_11_ALL.STORE" & _
              " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE"
    
    If chkSection.Value = 1 Then
        cString = cString & _
                  " LEFT JOIN FILE1_10SC ON FILE1_10.[SECTION] = FILE1_10SC.CODE"
    End If
                
    If chkGroup.Value = 1 Then
        cString = cString & _
                  " LEFT JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.CODE"
    End If
        
    If chkType.Value = 1 Then
        cString = cString & _
                  " LEFT JOIN FILE0_81 ON FILE0_81.CODE = FILE1_10.MODELSEX"
    End If
              
    If chkAge.Value = 1 Then
        cString = cString & _
                  " LEFT JOIN FILE0_82 ON FILE0_82.CODE = FILE1_10.MODELAGE"
    End If
    
    cString = cString & " WHERE file1_10.ISNOITEM = 0 "
                                                                          
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    If xMosm.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    If xmosm20.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm20.text)
    
    If XSECTION.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[SECTION] = " & (XSECTION.BoundText)
    If xSupp.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xStore.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " [store] = " & MyParn(xStore.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    If XOutlet.Value = 1 Then cWhere = cWhere & Tr(cWhere) & " Outlet = 1 "
    
    If xage.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If xsex.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)

    If IsDate(xDate2.Value) Then cWhere = cWhere & Tr(cWhere) & " [DATE] <= " & DateSq(xDate2.Value)
    
    If xall(0).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 1 "
    ElseIf xall(1).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 0 "
    End If
            
    If cWhere <> "" Then
        cString = cString & " AND " & cWhere
    End If
        
    cString = cString & " GROUP BY FILE1_11_ALL.STORE , STORE_ALL.DESCA"
        
    If chkSection.Value = 1 Then
        cString = cString & ",FILE1_10.[SECTION],FILE1_10SC.DESCA"
    End If
        
    If chkGroup.Value = 1 Then
        cString = cString & ",FILE1_10.[GROUP],FILE1_50.DESCA"
    End If
        
    If chkMosm.Value = 1 Then
        cString = cString & ",FILE1_10.MOSM"
    End If
    
    If chkType.Value = 1 Then
        cString = cString & ",FILE1_10.MODELSEX,FILE0_81.DESCA"
    End If
    
    If chkAge.Value = 1 Then
        cString = cString & ",FILE1_10.MODELAGE,FILE0_82.DESCA"
    End If
    
    cString = cString & " ORDER BY FILE1_11_ALL.STORE"
    
    If chkSection.Value = 1 Then
        cString = cString & ",FILE1_10.[SECTION]"
    End If
    
    If chkGroup.Value = 1 Then
        cString = cString & ",FILE1_10.[GROUP]"
    End If
    
    If chkMosm.Value = 1 Then
        cString = cString & ",FILE1_10.MOSM"
    End If
    
    If chkType.Value = 1 Then
        cString = cString & ",FILE1_10.MODELSEX"
    End If
    
    If chkAge.Value = 1 Then
        cString = cString & ",FILE1_10.MODELAGE"
    End If
    
    Set DATA16.Recordset = cmd(cString, con).Execute
End With
FIXGRID6
If GRID6.Rows > 0 Then GRID6.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub
Sub FIXGRID6()
Dim col As Long
Dim colRate As Long, ColSales As Long, ColNumbers As Long

With GRID6
    .RowHeight(0) = 1000
    .WordWrap = True
            
    .TextMatrix(0, col) = "«·›—⁄"
    .ColHidden(col) = True
    
    col = col + 1
    .TextMatrix(0, col) = "«·›—⁄"
    .ColWidth(col) = 2000
            
    If chkSection.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·ﬁ”„"
        .ColHidden(col) = True
        
        col = col + 1
        .TextMatrix(0, col) = "«·ﬁ”„"
        .ColWidth(col) = 2000
    End If
            
    If chkGroup.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·„Ã„Ê⁄…"
        .ColHidden(col) = True
        
        col = col + 1
        .TextMatrix(0, col) = "«·„Ã„Ê⁄…"
        .ColWidth(col) = 2000
    End If
            
    If chkMosm.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·„Ê”„"
        .ColWidth(col) = 1000
    End If
    
    If chkType.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·‰Ê⁄"
        .ColHidden(col) = True
            
        col = col + 1
        .TextMatrix(0, col) = "«·‰Ê⁄"
        .ColWidth(col) = 1000
    End If
    
    If chkAge.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·”‰"
        .ColHidden(col) = True
        
        col = col + 1
        .TextMatrix(0, col) = "«·”‰"
        .ColWidth(col) = 1200
    End If
     
     col = col + 1
     ColNumbers = col
    .TextMatrix(0, col) = "ﬂ„Ì… Ê«—œ"
    .ColWidth(col) = 1000
    .ColDataType(col) = flexDTDouble
    
    If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "Ê«—œ »”⁄— «· ﬂ·›…"
        .ColHidden(col) = Not bOpt10
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
    If bPrice Then
         col = col + 1
        .TextMatrix(0, col) = "Ê«—œ »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
     col = col + 1
    .TextMatrix(0, col) = "ﬂ„Ì… ’«œ—"
    .ColWidth(col) = 1000
    .ColDataType(col) = flexDTDouble

    If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "’«œ— »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
        .ColDataType(col) = flexDTDouble
    End If
    
    If bPrice Then
         col = col + 1
        .TextMatrix(0, col) = "’«œ— »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
     col = col + 1
     ColNet = col
    .TextMatrix(0, col) = "’«›Ï «·ﬂ„Ì…"
    .ColWidth(col) = 1000
    .ColDataType(col) = flexDTDouble
    
    If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "’«›Ï «·ﬂ„Ì… »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
        .ColDataType(col) = flexDTDouble
    End If
    
    
    If bPrice Then
         col = col + 1
        .TextMatrix(0, col) = "’«›Ï «·ﬂ„Ì… »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
     col = col + 1
     colRate = col
    .TextMatrix(0, col) = "‰”»… „»Ì⁄« "
    .ColWidth(col) = 1300
    .ColDataType(col) = flexDTDouble

    
     col = col + 1
     ColSales = col
    .TextMatrix(0, col) = "⁄œœ „»Ì⁄« "
    .ColWidth(col) = 1000
    .ColDataType(col) = flexDTDouble
     
     If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "„»Ì⁄«  »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
        .ColDataType(col) = flexDTDouble
    End If
    
    If bPrice Then
         col = col + 1
        .TextMatrix(0, col) = "„»Ì⁄«  »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
     col = col + 1
    .TextMatrix(0, col) = "—’Ìœ ⁄œœ"
    .ColWidth(col) = 1300
    .ColDataType(col) = flexDTDouble

    
    If bCost Then
        col = col + 1
        .TextMatrix(0, col) = "ﬁÌ„… «·—’Ìœ »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
        .ColDataType(col) = flexDTDouble
    End If
    
    If bPrice Then
        col = col + 1
        .TextMatrix(0, col) = "ﬁÌ„… «·—’Ìœ »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
    col = col + 1
    .TextMatrix(0, col) = "⁄œœ „ÊœÌ·« "
    .ColWidth(col) = 1300
    .ColDataType(col) = flexDTDouble

            
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    
    For Row = 1 To .Rows - 1
        If .ValueMatrix(Row, ColNet) <> 0 Then .TextMatrix(Row, colRate) = Round((.ValueMatrix(Row, ColSales) / .ValueMatrix(Row, ColNet)) * 100, 2)
    Next
    
    .SubtotalPosition = flexSTAbove
    For i = ColNumbers To .Cols - 1
        .Subtotal flexSTSum, -1, i, "#0", vbRed, vbYellow, True, "  "
    Next
    
    If .Rows > 1 Then
        If Val(.TextMatrix(1, ColNet)) <> 0 Then .TextMatrix(1, colRate) = Round((Val(.TextMatrix(1, ColSales)) / Val(.TextMatrix(1, ColNet))) * 100, 2)
    End If
    If .Rows > 1 Then
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End If
End With
End Sub
Private Sub myload7()
Dim cString  As String
Dim cWhere As String, cWhereField As String, cField As String
With GRID7
    If IsDate(xDate1.Value) Then
        cWhereField = " AND DATE >= " & DateSq(xDate1.Value)
    End If
        
    If chkSection.Value = 1 Then
        cField = cField & "," & "FILE1_10.[SECTION],FILE1_10SC.DESCA"
    End If
        
    If chkGroup.Value = 1 Then
        cField = cField & "," & "FILE1_10.[GROUP],FILE1_50.DESCA"
    End If
        
    If chkMosm.Value = 1 Then
        cField = cField & "," & "FILE1_10.MOSM"
    End If
    
    If chkType.Value = 1 Then
        cField = cField & "," & "FILE1_10.MODELSEX,FILE0_81.DESCA"
    End If
    
    If chkAge.Value = 1 Then
        cField = cField & "," & "FILE1_10.MODELAGE,FILE0_82.DESCA"
    End If
        
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "[IN]")
    
    If bCost Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "FILE1_11_ALL.[IN] * FILE1_11_ALL.COST")
    End If
    
    If bPrice Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "FILE1_11_ALL.[IN] * FILE1_11_ALL.PRICE_SALES")
    End If
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "[OUT] ")
    
    If bCost Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "FILE1_11_ALL.[OUT] * FILE1_11_ALL.COST")
    End If
    
    If bPrice Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "FILE1_11_ALL.[OUT] * FILE1_11_ALL.PRICE_SALES")
    End If
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) ")
    
    If bCost Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) * FILE1_11_ALL.COST")
    End If
    
    If bPrice Then
        cField = cField & "," & _
                myiif_r("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) * FILE1_11_ALL.PRICE_SALES")
    End If
    
    cField = cField & ",0," & _
            myiif("(type = '6')" & cWhereField, "([OUT])")
    
    If bCost Then
        cField = cField & "," & _
                myiif_r("(type = '6')" & cWhereField, "([OUT] * file1_11_all.cost) ")
    End If
    
    If bPrice Then
        cField = cField & "," & _
                myiif_r("(type = '6')" & cWhereField, "([OUT] * FILE1_11_ALL.PRICE_SALES) ")
    End If
    
    cField = cField & "," & _
            "SUM([IN]-[OUT])"
    
    If bCost Then
        cField = cField & "," & _
                 " ROUND(SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ),2)"
    End If
    
    If bPrice Then
        cField = cField & "," & _
                " ROUND(SUM(([IN]-[OUT]) * FILE1_11_ALL.PRICE_SALES ),2)"
    End If
    
    cString = "SELECT FILE4_10.DESCA, FACT.DESCA, FILE1_10.GRMODEL , max(FILE1_10.GRDESCA) " & _
              cField & _
              ", COUNT(DISTINCT FILE1_10.MODEL) " & _
              " FROM  FILE1_10 INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
              " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
              " INNER JOIN FILE4_10 ON FILE1_10.code = FILE4_10.CODE" & _
              " INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE"
    
    If chkSection.Value = 1 Then
        cString = cString & _
                  " LEFT JOIN FILE1_10SC ON FILE1_10.[SECTION] = FILE1_10SC.CODE"
    End If
        
    If chkGroup.Value = 1 Then
        cString = cString & _
                  " LEFT JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.CODE"
    End If
        
    If chkType.Value = 1 Then
        cString = cString & _
                  " LEFT JOIN FILE0_81 ON FILE0_81.CODE = FILE1_10.MODELSEX"
    End If
              
    If chkAge.Value = 1 Then
        cString = cString & _
                  " LEFT JOIN FILE0_82 ON FILE0_82.CODE = FILE1_10.MODELAGE"
    End If
    
    cString = cString & " WHERE file1_10.ISNOITEM = 0 "
                                                                          
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    If xMosm.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    If xmosm20.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm20.text)
    
    If XSECTION.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[SECTION] = " & (XSECTION.BoundText)
    If xSupp.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xStore.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " [store] = " & MyParn(xStore.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    If XOutlet.Value = 1 Then cWhere = cWhere & Tr(cWhere) & " Outlet = 1 "
    
    If xage.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If xsex.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)

    If IsDate(xDate2.Value) Then cWhere = cWhere & Tr(cWhere) & " [DATE] <= " & DateSq(xDate2.Value)
    
    If xall(0).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 1 "
    ElseIf xall(1).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 0 "
    End If
            
    If cWhere <> "" Then
        cString = cString & " AND " & cWhere
    End If
        
    cString = cString & " GROUP BY FILE4_10.CODE, FILE4_10.DESCA, FACT.DESCA,FILE1_10.GRMODEL"
        
    If chkSection.Value = 1 Then
        cString = cString & "," & "FILE1_10.[SECTION]" & "," & "FILE1_10SC.DESCA"
    End If
        
    If chkGroup.Value = 1 Then
        cString = cString & "," & "FILE1_10.[GROUP]" & "," & "FILE1_50.DESCA"
    End If
        
    If chkMosm.Value = 1 Then
        cString = cString & "," & "FILE1_10.MOSM"
    End If
    
    If chkType.Value = 1 Then
        cString = cString & "," & "FILE1_10.MODELSEX,FILE0_81.DESCA"
    End If
    
    If chkAge.Value = 1 Then
        cString = cString & "," & "FILE1_10.MODELAGE,FILE0_82.DESCA"
    End If
    
    cString = cString & " ORDER BY FILE4_10.CODE, FILE4_10.DESCA, FACT.DESCA"
    
    If chkSection.Value = 1 Then
        cString = cString & "," & "FILE1_10.[SECTION]"
    End If
    
    If chkGroup.Value = 1 Then
        cString = cString & "," & "FILE1_10.[GROUP]"
    End If
    
    If chkMosm.Value = 1 Then
        cString = cString & "," & "FILE1_10.MOSM"
    End If
    
    If chkType.Value = 1 Then
        cString = cString & "," & "FILE1_10.MODELSEX"
    End If
    
    If chkAge.Value = 1 Then
        cString = cString & "," & "FILE1_10.MODELAGE"
    End If
    
    Set DATA17.Recordset = cmd(cString, con).Execute
End With
FIXGRID7
If GRID7.Rows > 0 Then GRID7.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub
Sub FIXGRID7()
Dim col As Long
Dim colRate As Long, ColSales As Long, ColNumbers As Long

With GRID7
    .RowHeight(0) = 1000
    .WordWrap = True
                            
    .TextMatrix(0, col) = "«·„Ê—œ"
    .ColWidth(col) = 2000
    
    col = col + 1
    .TextMatrix(0, col) = "«·„’‰⁄"
    .ColWidth(col) = 2000
            
    col = col + 1
    .TextMatrix(0, col) = "„Ã„Ê⁄… «·„ÊœÌ·"
    .ColWidth(col) = 2000
            
    col = col + 1
    .TextMatrix(0, col) = "»Ì«‰"
    .ColWidth(col) = 2000
            
    If chkSection.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·ﬁ”„"
        .ColHidden(col) = True
        
        col = col + 1
        .TextMatrix(0, col) = "«·ﬁ”„"
        .ColWidth(col) = 2000
    End If
            
    If chkGroup.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·„Ã„Ê⁄…"
        .ColHidden(col) = True
        
        col = col + 1
        .TextMatrix(0, col) = "«·„Ã„Ê⁄…"
        .ColWidth(col) = 2000
    End If
            
    If chkMosm.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·„Ê”„"
        .ColWidth(col) = 1000
    End If
    
    If chkType.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·‰Ê⁄"
        .ColHidden(col) = True
            
        col = col + 1
        .TextMatrix(0, col) = "«·‰Ê⁄"
        .ColWidth(col) = 1000
    End If
    
    If chkAge.Value = 1 Then
        col = col + 1
        .TextMatrix(0, col) = "«·”‰"
        .ColHidden(col) = True
        
        col = col + 1
        .TextMatrix(0, col) = "«·”‰"
        .ColWidth(col) = 1200
    End If
     
     col = col + 1
     ColNumbers = col
    .TextMatrix(0, col) = "ﬂ„Ì… Ê«—œ"
    .ColWidth(col) = 1000
    .ColDataType(col) = flexDTDouble
    
     If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "Ê«—œ »”⁄— «· ﬂ·›…"
        .ColHidden(col) = Not bOpt10
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
    If bPrice Then
         col = col + 1
        .TextMatrix(0, col) = "Ê«—œ »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
     col = col + 1
    .TextMatrix(0, col) = "ﬂ„Ì… ’«œ—"
    .ColWidth(col) = 1000
    .ColDataType(col) = flexDTDouble

    If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "’«œ— »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
        .ColDataType(col) = flexDTDouble
    End If
    
    If bPrice Then
         col = col + 1
        .TextMatrix(0, col) = "’«œ— »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
     col = col + 1
     ColNet = col
    .TextMatrix(0, col) = "’«›Ï «·ﬂ„Ì…"
    .ColWidth(col) = 1000
    .ColDataType(col) = flexDTDouble
    
    If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "’«›Ï «·ﬂ„Ì… »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
        .ColDataType(col) = flexDTDouble
    End If
    
     If bPrice Then
         col = col + 1
        .TextMatrix(0, col) = "’«›Ï «·ﬂ„Ì… »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
     col = col + 1
     colRate = col
    .TextMatrix(0, col) = "‰”»… „»Ì⁄« "
    .ColWidth(col) = 1300
    .ColDataType(col) = flexDTDouble

    
     col = col + 1
     ColSales = col
    .TextMatrix(0, col) = "⁄œœ „»Ì⁄« "
    .ColWidth(col) = 1000
    .ColDataType(col) = flexDTDouble
     
     If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "„»Ì⁄«  »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
        .ColDataType(col) = flexDTDouble
    End If
    
    If bPrice Then
         col = col + 1
        .TextMatrix(0, col) = "„»Ì⁄«  »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
     col = col + 1
    .TextMatrix(0, col) = "—’Ìœ ⁄œœ"
    .ColWidth(col) = 1300
    .ColDataType(col) = flexDTDouble

    If bCost Then
        col = col + 1
        .TextMatrix(0, col) = "ﬁÌ„… «·—’Ìœ »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
        .ColDataType(col) = flexDTDouble
    End If
    
    If bPrice Then
        col = col + 1
        .TextMatrix(0, col) = "ﬁÌ„… «·—’Ìœ »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
        .ColDataType(col) = flexDTDouble
    End If
    
    col = col + 1
    .TextMatrix(0, col) = "⁄œœ „ÊœÌ·« "
    .ColWidth(col) = 1300
    .ColDataType(col) = flexDTDouble
            
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    
    For Row = 1 To .Rows - 1
        If .ValueMatrix(Row, ColNet) <> 0 Then .TextMatrix(Row, colRate) = Round((.ValueMatrix(Row, ColSales) / .ValueMatrix(Row, ColNet)) * 100, 2)
    Next
    
    .SubtotalPosition = flexSTAbove
    For i = ColNumbers To .Cols - 1
        .Subtotal flexSTSum, -1, i, "#0", vbRed, vbYellow, True, "  "
    Next
    
    If .Rows > 1 Then
        If Val(.TextMatrix(1, ColNet)) <> 0 Then .TextMatrix(1, colRate) = Round((Val(.TextMatrix(1, ColSales)) / Val(.TextMatrix(1, ColNet))) * 100, 2)
    End If
    If .Rows > 1 Then
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End If
End With
End Sub
Private Sub myload5()
Dim cString  As String
Dim cField As String, cWhere As String, cWhereField As String
With GRID5
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )", "([IN]) ")
    
    If bCost Then
        cField = cField & "," & _
                myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )", "(FILE1_11_ALL.[IN] * FILE1_11_ALL.COST)")
    End If
    
    If bPrice Then
        cField = cField & "," & _
                myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )", "(FILE1_11_ALL.[IN] * FILE1_11_ALL.PRICE_SALES)")
    End If
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )", "([OUT])")
    
    If bCost Then
        cField = cField & "," & _
                myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )", "(FILE1_11_ALL.[OUT] * FILE1_11_ALL.COST)")
    End If
    
    If bPrice Then
        cField = cField & "," & _
                myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )", "(FILE1_11_ALL.[OUT] * FILE1_11_ALL.PRICE_SALES)")
    End If
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )", "([IN]-[OUT]) ")
    
    If bCost Then
        cField = cField & "," & _
                myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )", "(([IN]-[OUT]) * FILE1_11_ALL.COST)")
    End If
    
    If bPrice Then
        cField = cField & "," & _
                myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )", "(([IN]-[OUT]) *  FILE1_11_ALL.PRICE_SALES) ")
    End If
    
    cField = cField & "," & _
            myiif("type = '6' ", " [OUT] ")
    
    cField = cField & "," & _
            myiif("type = '6' ", " ([OUT]* FILE1_11_ALL.COST) ")
    
    cField = cField & "," & _
            myiif("type = '6' ", " (FILE1_11_ALL.TOTAL ) ")
    
    cField = cField & "," & _
            myiif("type = '6' ", " FILE1_11_ALL.TOTAL -([OUT]* FILE1_11_ALL.COST) ")
                                
    
    cField = cField & "," & _
            "SUM([IN]-[OUT])"
    
    If bCost Then
        cField = cField & "," & _
                 " ROUND(SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ),2)"
    End If
    
    If bPrice Then
        cField = cField & "," & _
                " ROUND(SUM(([IN]-[OUT]) * FILE1_11_ALL.PRICE_SALES ),2)"
    End If
                                        
    cString = " SELECT  FILE4_10.DESCA AS supldesca, FACT.DESCA AS factdesca, FILE1_10.MOSM  ,  FILE1_10.MODELFACT0, FILE1_10.desca , STORE_ALL.DESCA  " & _
                cField & _
                " , file0_81.desca , file0_82.desca , file1_10SC.DESCA , FILE1_50.DESCA , GrModel ,MOSM2     " & _
                "  FROM    FILE1_10 INNER JOIN FILE4_10 ON FILE1_10.code = FILE4_10.CODE " & _
                " INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE" & _
                " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
                " INNER JOIN STORE_ALL ON STORE_ALL.CODE = FILE1_11_ALL.STORE " & _
                " LEFT JOIN FILE0_81 ON FILE0_81.CODE = FILE1_10.ModelSex" & _
                " LEFT JOIN FILE0_82 ON FILE0_82.CODE = FILE1_10.ModelAge" & _
                " LEFT JOIN FILE1_10SC ON FILE1_10SC.CODE = FILE1_10.[SECTION]" & _
                " LEFT JOIN FILE1_50 ON FILE1_50.CODE = FILE1_10.[GROUP]" & _
                " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE"
    
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    If xMosm.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    
    If xmosm20.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm20.text)
    
    If XSECTION.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[SECTION] = " & (XSECTION.BoundText)
    If xSupp.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xStore.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " [store] = " & MyParn(xStore.BoundText)
    
    If Not bOpt5 Then
        cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    End If
    
    If XOutlet.Value = 1 Then cWhere = cWhere & Tr(cWhere) & " Outlet = 1 "
    
    If xage.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If xsex.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)


    If IsDate(xDate1.Value) Then cWhere = cWhere & Tr(cWhere) & " [DATE] >= " & DateSq(xDate1.Value)
    If IsDate(xDate2.Value) Then cWhere = cWhere & Tr(cWhere) & " [DATE] <= " & DateSq(xDate2.Value)
    
    If xall(0).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 1 "
    ElseIf xall(1).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 0 "
    End If
    
    
    If cWhere <> "" Then
        cString = cString & " Where " & cWhere
    End If
    cString = cString & " GROUP BY FILE4_10.CODE, FILE4_10.DESCA,STORE_ALL.DESCA,STORE_ALL.CODE , FACT.DESCA, FILE1_10.MOSM, FILE1_10.MODELFACT0, FILE1_10.desca , FILE1_10.MODEL  , FILE1_10.desca , FILE1_10.MODEL   , file0_81.desca , file0_82.desca , file1_10SC.DESCA , FILE1_50.DESCA , GrModel ,MOSM2"
    cString = cString & " ORDER BY FILE4_10.CODE, FILE4_10.DESCA, FACT.DESCA, FILE1_10.MOSM, FILE1_10.MODELFACT0 ,STORE_ALL.CODE "
    
    If Check1.Value = 1 Then
        textquery.text = cString
    Else
        Set DATA15.Recordset = cmd(cString, con).Execute
    End If
       
End With
FIXGRID5
If GRID5.Rows > 0 Then GRID5.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub
Sub FIXGRID5()
With GRID5
Dim ColNumbers As Long
    
    .RowHeight(0) = 800
    .WordWrap = True
    
    .TextMatrix(0, 0) = "«·„Ê—œ"
    .ColWidth(0) = 1200
    
    .TextMatrix(0, 1) = "«·„’‰⁄"
    .ColWidth(1) = 1200
    
    .TextMatrix(0, 2) = "„Ê”„"
    .ColWidth(2) = 1200
    
    .TextMatrix(0, 3) = "„ÊœÌ·"
    .ColWidth(3) = 1200
    
    .TextMatrix(0, 4) = "«·’‰›"
    .ColWidth(4) = 1200
    
    .TextMatrix(0, 5) = "«·›—⁄"
    .ColWidth(5) = 1200
    
    .TextMatrix(0, 6) = "ﬂ„Ì… Ê«—œ"
    .ColWidth(6) = 1200

    ColNumbers = 6
    col = 6
    
    If bCost Then
         col = col + 1
        .TextMatrix(0, col) = "‰ﬂ·›… «·Ê«—œ"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
    End If
    
    If bPrice Then
        col = col + 1
        .TextMatrix(0, col) = " ﬁÌ„ «·Ê«—œ „” Â·ﬂ"
        .ColWidth(col) = 1300
    End If
    
     col = col + 1
    .TextMatrix(0, col) = "ﬂ„Ì… ’«œ—"
    .ColWidth(col) = 800
               
    If bCost Then
         col = col + 1
        .TextMatrix(0, col) = " ﬂ·›… ’«œ—"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
    End If
           
    If bPrice Then
        col = col + 1
        .TextMatrix(0, col) = " ﬁÌ„ ’«œ— „” Â·ﬂ"
        .ColWidth(col) = 1300
    End If
               
    col = col + 1
    .TextMatrix(0, col) = "’«›Ï «·ﬂ„Ì…"
    .ColWidth(col) = 800
    
    If bCost Then
        col = col + 1
        .TextMatrix(0, col) = "’«›Ï  »”⁄— «· ﬂ·›…"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
    End If
    
    If bPrice Then
        col = col + 1
        .TextMatrix(0, col) = "’«›Ì  »”⁄— «·„” Â·ﬂ"
        .ColWidth(col) = 1300
    End If
    
     col = col + 1
    .TextMatrix(0, col) = "ﬂ„Ì… „»Ì⁄« "
    .ColWidth(col) = 800
        
    If bCost Then
         col = col + 1
        .TextMatrix(0, col) = " ﬂ·›… „»Ì⁄« "
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
    End If
    
    col = col + 1
    .TextMatrix(0, col) = "ﬁÌ„… „»Ì⁄« "
    .ColWidth(col) = 1300
    
    col = col + 1
    .TextMatrix(0, col) = "—»Õ „»Ì⁄« "
    .ColWidth(col) = 1300
    
    col = col + 1
    .TextMatrix(0, col) = "«·—’Ìœ"
    .ColWidth(col) = 800
    
    If bCost Then
         col = col + 1
        .TextMatrix(0, col) = " ﬂ·›… «·—’Ìœ"
        .ColWidth(col) = 1300
        .ColHidden(col) = Not bOpt10
    End If
    
    If bPrice Then
        col = col + 1
        .TextMatrix(0, col) = " ﬁÌ„ —’Ìœ »«·„” Â·ﬂ"
        .ColWidth(col) = 1300
    End If
    
    col = col + 1
    .TextMatrix(0, col) = "«·‰Ê⁄"
    .ColWidth(col) = 1200
    
    col = col + 1
    .TextMatrix(0, col) = "«·›∆… «·⁄„—Ì…"
    .ColWidth(col) = 1200
    
    col = col + 1
    .TextMatrix(0, col) = "«·ﬁ”„"
    .ColWidth(col) = 1200
    
    col = col + 1
    .TextMatrix(0, col) = "„Ã„Ê⁄… «·„ÊœÌ·« "
    .ColWidth(col) = 1200
    
    col = col + 1
    .TextMatrix(0, col) = "„Ã„Ê⁄… „ÊœÌ·"
    .ColWidth(col) = 1200
    
    col = col + 1
    .TextMatrix(0, col) = "„Ê”„ œ«Œ·Ï"
    .ColWidth(col) = 1200
    
    
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 6) = 4
            
    .SubtotalPosition = flexSTAbove
    
    For i = ColNumbers To .Cols - 7
        .Subtotal flexSTSum, -1, i, "#0", vbRed, vbYellow, True, "  "
    Next i
            
    If .Rows > 1 Then
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End If
End With
End Sub

Private Sub textquery_GotFocus()
textquery.SelStart = 0
textquery.SelLength = Len(textquery.text)
End Sub

