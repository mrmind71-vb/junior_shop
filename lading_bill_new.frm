VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form ladingBillfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "»Ê«·’ ‘Õ‰ „”»ﬁ…"
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20370
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
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   11055
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1995
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   1995
      ScaleWidth      =   20370
      TabIndex        =   17
      Top             =   8490
      Width           =   20370
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   11790
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   0
         Width           =   5415
         Begin VB.TextBox xdate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   900
            MaxLength       =   300
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   42
            Tag             =   "filter"
            Top             =   225
            Width           =   1365
         End
         Begin VB.CheckBox chkDay 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "«·ÌÊ„"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   270
            Width           =   690
         End
         Begin VB.TextBox xdate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2880
            MaxLength       =   300
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   40
            Tag             =   "filter"
            Top             =   225
            Width           =   1275
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Õ Ì"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   2385
            TabIndex        =   44
            Tag             =   "COLOR"
            Top             =   270
            Width           =   420
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   " «—ÌŒ »Ê·Ì’…"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   15
            Left            =   4230
            TabIndex        =   43
            Tag             =   "COLOR"
            Top             =   270
            Width           =   1050
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   495
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   0
         Width           =   11265
         Begin VB.TextBox xOnline_doc 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2565
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   225
            Width           =   1635
         End
         Begin VB.TextBox xdoc_no_ret 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   225
            Width           =   1680
         End
         Begin VB.TextBox xship_no 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   8460
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   225
            Width           =   1590
         End
         Begin VB.TextBox xdoc_no 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   5355
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   225
            Width           =   1635
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "—ﬁ„ «·ÿ·»Ì…"
            Height          =   240
            Left            =   4275
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   270
            Width           =   1005
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "«·„— Ã⁄"
            Height          =   285
            Left            =   1845
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   270
            Width           =   780
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "—ﬁ„ «·»Ê·Ì’…"
            Height          =   240
            Index           =   0
            Left            =   10125
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   270
            Width           =   1050
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "—ﬁ„ «·›« Ê—…"
            Height          =   240
            Left            =   7065
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   270
            Width           =   1005
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   495
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   585
         Width           =   5820
         Begin VB.TextBox xdate_Return2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1125
            MaxLength       =   300
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Tag             =   "filter"
            Top             =   225
            Width           =   1230
         End
         Begin VB.CheckBox chkday_Return 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "«·ÌÊ„"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   270
            Width           =   780
         End
         Begin VB.TextBox xdate_return1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3375
            MaxLength       =   300
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            Tag             =   "filter"
            Top             =   225
            Width           =   1185
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Õ Ì"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   4
            Left            =   2520
            TabIndex        =   32
            Tag             =   "COLOR"
            Top             =   270
            Width           =   375
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   " «—ÌŒ  „— Ã⁄"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   5
            Left            =   4725
            TabIndex        =   31
            Tag             =   "COLOR"
            Top             =   270
            Width           =   1005
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   6345
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   585
         Width           =   5415
         Begin VB.TextBox xdate_Delivery1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2880
            MaxLength       =   300
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   36
            Tag             =   "filter"
            Top             =   225
            Width           =   1275
         End
         Begin VB.CheckBox chkDay_delivery 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "«·ÌÊ„"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   270
            Width           =   690
         End
         Begin VB.TextBox xdate_Delivery2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   945
            MaxLength       =   300
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            Tag             =   "filter"
            Top             =   225
            Width           =   1320
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   " «—ÌŒ «” ·«„"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   3
            Left            =   4275
            TabIndex        =   38
            Tag             =   "COLOR"
            Top             =   270
            Width           =   915
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Õ Ì"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   2340
            TabIndex        =   37
            Tag             =   "COLOR"
            Top             =   270
            Width           =   420
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   11790
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   585
         Width           =   5415
         Begin VB.TextBox xdate_mail1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2880
            MaxLength       =   300
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   48
            Tag             =   "filter"
            Top             =   225
            Width           =   1275
         End
         Begin VB.CheckBox chkdayMail 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "«·ÌÊ„"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   270
            Width           =   690
         End
         Begin VB.TextBox xdate_mail2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   945
            MaxLength       =   300
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   46
            Tag             =   "filter"
            Top             =   225
            Width           =   1320
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   " «—ÌŒ «Ì„Ì·"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   6
            Left            =   4275
            TabIndex        =   50
            Tag             =   "COLOR"
            Top             =   270
            Width           =   1005
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Õ Ì"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   7
            Left            =   2430
            TabIndex        =   49
            Tag             =   "COLOR"
            Top             =   270
            Width           =   375
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   3105
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   1215
         Width           =   6135
         Begin Threed.SSCommand cmdShip 
            Height          =   465
            Left            =   3960
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   180
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   820
            _Version        =   196610
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
            Caption         =   "‘—ﬂ… «·‘Õ‰"
            TagVariant      =   "‘—ﬂ… «·‘Õ‰"
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdType 
            Height          =   465
            Left            =   1890
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   180
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   820
            _Version        =   196610
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
            Caption         =   "‰Ê⁄ «·»Ê·Ì’…"
            TagVariant      =   "‰Ê⁄ «·»Ê·Ì’…"
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdExcel 
            Height          =   465
            Left            =   45
            TabIndex        =   75
            Top             =   180
            Width           =   1815
            _ExtentX        =   3201
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
            Picture         =   "lading_bill_new.frx":0000
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            ShapeSize       =   1
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   9315
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   1215
         Width           =   7260
         Begin VB.TextBox xPhone 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   225
            Width           =   2445
         End
         Begin VB.TextBox xName 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3780
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   225
            Width           =   2490
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFFF&
            Caption         =   " ·Ì›Ê‰ «·⁄„Ì·"
            Height          =   240
            Left            =   2655
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   270
            Width           =   1005
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "≈”„ «·⁄„Ì·"
            Height          =   240
            Left            =   6390
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   270
            Width           =   780
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   495
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   1260
         Width           =   2580
         Begin Threed.SSCommand cmdExit 
            Height          =   465
            Left            =   45
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   135
            Width           =   1230
            _ExtentX        =   2170
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
            Picture         =   "lading_bill_new.frx":2414
            Caption         =   "///"
            Alignment       =   8
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            ShapeSize       =   1
         End
         Begin Threed.SSCommand cmdGo 
            Height          =   465
            Left            =   1305
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   135
            Width           =   1230
            _ExtentX        =   2170
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
            Picture         =   "lading_bill_new.frx":4737
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            ShapeSize       =   1
         End
      End
      Begin VB.Frame Frame8 
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
         Height          =   645
         Left            =   16605
         TabIndex        =   59
         Top             =   1260
         Width           =   3615
         Begin MSDataListLib.DataCombo xYear 
            Height          =   330
            Left            =   1665
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   225
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo xMonth 
            Height          =   330
            Left            =   90
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   225
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "«·”‰…"
            Height          =   285
            Left            =   2970
            TabIndex        =   63
            Tag             =   "Color"
            Top             =   270
            Width           =   510
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "«·‘Â—"
            Height          =   330
            Left            =   855
            TabIndex        =   62
            Tag             =   "Color"
            Top             =   270
            Width           =   555
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   17235
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   585
         Width           =   2985
         Begin VB.OptionButton optMail 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "«·ﬂ·"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   2205
            RightToLeft     =   -1  'True
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   270
            Value           =   -1  'True
            Width           =   645
         End
         Begin VB.OptionButton optMail 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "» «—ÌŒ «Ì„Ì·"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   900
            RightToLeft     =   -1  'True
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   270
            Width           =   1230
         End
         Begin VB.OptionButton optMail 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "»œÊ‰ "
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   270
            Width           =   735
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   17235
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   0
         Width           =   2985
         Begin VB.OptionButton optClose 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "„€·ﬁ…"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   270
            Width           =   780
         End
         Begin VB.OptionButton optClose 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "€Ì— „€·ﬁ…"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   1035
            RightToLeft     =   -1  'True
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   270
            Value           =   -1  'True
            Width           =   1050
         End
         Begin VB.OptionButton optClose 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "«·ﬂ·"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   2205
            RightToLeft     =   -1  'True
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   270
            Width           =   645
         End
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   8205
      Left            =   0
      TabIndex        =   0
      Top             =   90
      Width           =   20175
      _cx             =   35586
      _cy             =   14473
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   8454143
      ForeColorSel    =   128
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   15
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
      FrozenCols      =   2
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSAdodcLib.Adodc data11 
      Height          =   330
      Left            =   -1575
      Top             =   3015
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
      Left            =   4455
      Top             =   1575
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   2475
      Top             =   1125
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   270
      Top             =   1350
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   10680
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   661
      _Version        =   196610
      BackColor       =   16777215
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
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
         Left            =   4095
         TabIndex        =   3
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
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
         Left            =   8100
         TabIndex        =   4
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
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
         Index           =   3
         Left            =   12150
         TabIndex        =   5
         Top             =   45
         Width           =   3960
         _ExtentX        =   6985
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
         Index           =   4
         Left            =   16155
         TabIndex        =   6
         Top             =   45
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2850
      Left            =   1035
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   3825
      Visible         =   0   'False
      Width           =   6225
      Begin VB.Frame Frame11 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   0
         Width           =   4740
         Begin VB.OptionButton optDelivery 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "»œÊ‰  «—ÌŒ «” ·«„"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   1935
            RightToLeft     =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   270
            Width           =   1725
         End
         Begin VB.OptionButton optDelivery 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "» «—ÌŒ «” ·«„"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   135
            RightToLeft     =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   270
            Width           =   1410
         End
         Begin VB.OptionButton optDelivery 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "«·ﬂ·"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   3825
            RightToLeft     =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   270
            Value           =   -1  'True
            Width           =   780
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   855
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1260
         Visible         =   0   'False
         Width           =   3615
         Begin VB.OptionButton optDoc 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "»›« Ê—…"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   270
            Width           =   870
         End
         Begin VB.OptionButton optDoc 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "»œÊ‰ ›« Ê—…"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   1305
            RightToLeft     =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   270
            Width           =   1230
         End
         Begin VB.OptionButton optDoc 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "«·ﬂ·"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   2700
            RightToLeft     =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   270
            Value           =   -1  'True
            Width           =   780
         End
      End
   End
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   16
      Top             =   10485
      Visible         =   0   'False
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "ladingBillfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bEdit As Boolean
Public myForm As Form
Public sFilter As String, sSup As String
Dim bStopCell As Boolean
'Dim rsShip As New ADODB.Recordset
Dim bIgClick As Boolean
Dim cList1 As String, cList2 As String
Dim oSearchInv As New Search_abd
Dim oSearchship As New Search_abd
Dim oSearchLading As New Search_abd
Dim con As New ADODB.Connection
Private Sub myLoadGrd()
Dim cString As String
Dim cWhere As String

Dim aFields(15 + 1)
aFields(0) = "LADING_BILL.TYPE AS [‰Ê⁄ «·»Ê·Ì’…]"
aFields(1) = "LADING_BILL.SHIP_NO AS [—ﬁ„ «·»Ê·Ì’…]"
aFields(2) = "FORMAT(LADING_BILL.DATE,'yyyy/M/d') AS [ «—ÌŒ «·»Ê·Ì’…]"
aFields(3) = "LADING_BILL.SHIP AS [‘—ﬂ… «·‘Õ‰]"
aFields(4) = "LADING_BILL.CHARGE1 AS [„’«—Ì› ‘Õ‰]"
aFields(5) = "LADING_BILL.CHARGE2 AS [„’«—Ì› »Ê·Ì’…]"
aFields(6) = "FILE6_20H.ONLINE_DOC AS [—ﬁ„ «·ÿ·»Ì…]"
aFields(7) = "LADING_BILL.DOC_NO AS [—ﬁ„ «·›« Ê—…]"
aFields(8) = "FORMAT(FILE6_20h.DATE,'yyyy/M/d') AS [ «—ÌŒ «·›« Ê—…]"
aFields(9) = "FILE6_20H.NAME AS [«”„ «·⁄„Ì·]"
aFields(10) = "FILE6_20H.PHONE AS [ ·Ì›Ê‰ «·⁄„Ì·]"
aFields(11) = "FILE6_20H.ADDRESS AS [«·⁄‰Ê«‰]"
aFields(11 + 1) = "FORMAT(LADING_BILL.DATE_MAIL,'yyyy/M/d') AS [ «—ÌŒ «·«Ì„Ì·]"
aFields(12 + 1) = "LADING_BILL.DOC_NO_RET AS [›« Ê—… «·„— Ã⁄]"
aFields(13 + 1) = "FORMAT(F.DATE,'yyyy/M/d') AS [ «—ÌŒ «·„— Ã⁄]"
aFields(14 + 1) = "FORMAT(LADING_BILL.DATE_DELIVERY,'yyyy/M/d') AS [ «—ÌŒ «·«” ·«„]"
aFields(15 + 1) = "LADING_BILL.ID"

cString = "SELECT " & arString(aFields, ",") & _
          " FROM LADING_BILL " & _
          " INNER JOIN FILE6_20H ON LADING_BILL.DOC_NO = FILE6_20H.DOC_NO" & _
          " LEFT JOIN FILE6_20H AS F ON LADING_BILL.DOC_NO_RET = F.DOC_NO"

cWhere = retWhere

If cWhere <> "" Then
    cString = cString & " WHERE " & cWhere
End If

cString = cString & " ORDER BY LADING_BILL.[DATE],LADING_BILL.ID"
Set data11.Recordset = mycmd(cString, con)

myAddItem
fixGrd
CalcTotals
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub myload()
myLoadGrd
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
End Sub

Private Sub chkDay_Click()
If bIgClick Then Exit Sub
myLoadGrd
End Sub
Private Sub chkDuplicate_Click()
myload
End Sub

Private Sub chkNoNum_Click()
myload
End Sub

Private Sub chkNoQuant_Click()
myload
End Sub

Private Sub chkOrderReceipt_Click()
myload
End Sub
Private Sub chkDay_delivery_Click()
If bIgClick Then Exit Sub
myLoadGrd
End Sub

Private Sub chkday_Return_Click()
If bIgClick Then Exit Sub
myLoadGrd
End Sub

Private Sub chkdaymail_Click()
If bIgClick Then Exit Sub
myLoadGrd
End Sub

Private Sub cmdExcel_Click()
Me.MousePointer = 11
Dim i As Long
Dim aRow As Variant
If grid1.Rows > 1 Then
    aSub = AddFlag(Empty, "row", 1)
    aSub = AddFlag(aSub, "bold", True)
    aSub = AddFlag(aSub, "word_wrap", False)
    aSub = AddFlag(aSub, "back_color", 19)
    aRow = AddFlag(aRow, aSub)
End If
ToFileExelNew grid1, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption)
Me.MousePointer = 0
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdGo_Click()
myload
End Sub

Private Sub cmdShip_Click()
ShipLookup Me, oSearchship, , , IIf(cmdShip.Tag <> "", "ﬂ· «·‘—ﬂ« ", "")
End Sub
Private Sub cmdType_Click()
LadingLookup Me, oSearchLading, , , IIf(cmdType.Tag <> "", "ﬂ· «‰Ê«⁄ «·»Ê«·’", "")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo) And ActiveControl.Tag <> "filter" Then
        KeyCode = 0
        grid1.SetFocus
    End If
End If
End Sub
Private Sub Form_Resize()
Dim nHeight As Long
nHeight = Me.Height - (grid1.Top + Picture1.Height + panel1(0).Height + 700)
grid1.Height = IIf(nHeight < 2000, 2000, nHeight)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'SaveText Me
Set ladingBillfrm = Nothing
Err.Clear
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal col As Long)
If Not validRow(Row) Then Exit Sub

If grid1.Row = grid1.Rows - 1 Then
    myAddItem
ElseIf Row = grid1.Rows - 2 And (col = 2) Then
    'MyEditItem grid1, Row, Col
End If

Dim bNew As Boolean
bNew = grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = ""

Dim nId As Variant
If myreplace(Row, nId) Then
    If bNew Then
        If IsNull(nId) Then
            myLoadGrd
        Else
             grid1.TextMatrix(Row, grid1.Cols - 1) = nId
             'grid1.TextMatrix(Row, grid1.Cols - 2) = GetComputerName
        End If
        grid1.ShowCell grid1.Rows - 1, 1
    End If
    grid1_EnterCell
Else
    myLoadGrd
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myLoadGrd
End Sub
Private Sub CalcTotal(Row As Long)
End Sub
Private Sub grid1_DblClick()
If grid1.Row > 0 And grid1.col = 1 And grid1.Editable = flexEDNone Then
    sShip_no = GetNumbersFromString(InputBox("enter Ship_No", "Change ship No"))
    If Trim(sShip_no) = "" Then Exit Sub
    
    con.BeginTrans
    On Error GoTo myerror
    con.Execute "update LADING_BILL SET " & _
                 "LADING_BILL.SHIP_NO = " & MyParn(sShip_no) & _
                 " FROM LADING_BILL WHERE ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
    
    con.Execute "update file6_20h " & _
                " set file6_20h.SHIP_NO = " & MyParn(sShip_no) & _
                " FROM FILE6_20H WHERE DOC_NO = " & MyParn(grid1.TextMatrix(grid1.Row, 13)), nRecords
    con.CommitTrans
    
    If nRecords > 0 Then Inform " „  ⁄œÌ· " & nRecords & " ”Ã· »‰Ã«Õ"
    myLoadGrd
ElseIf grid1.Row > 0 And grid1.col = 7 And grid1.Editable = flexEDNone Then
    sales_onlinefrm.bEdit = True
    sales_onlinefrm.sDoc_no = grid1.TextMatrix(grid1.Row, 7)
    sales_onlinefrm.Show
ElseIf grid1.Row > 0 And grid1.col = 13 And grid1.Editable = flexEDNone Then
    sales_onlinefrm.bEdit = True
    sales_onlinefrm.sDoc_no = grid1.TextMatrix(grid1.Row, 13)
    sales_onlinefrm.Show
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
If TransCount(con) > 0 Then con.RollbackTrans
End Sub

Private Sub grid1_EnterCell()

'aFields(0) = "LADING_BILL.SHIP_NO AS [—ﬁ„ «·»Ê·Ì’…]"
'aFields(1) = "FORMAT(LADING_BILL.DATE,'yyyy/M/d') AS [ «—ÌŒ «·»Ê·Ì’…]"
'aFields(2) = "LADING_BILL.SHIP AS [‘—ﬂ… «·‘Õ‰]"
'aFields(3) = "LADING_BILL.TYPE AS [‰Ê⁄ «·»Ê·Ì’…]"
'aFields(4) = "LADING_BILL.CHARGE1 AS [„’«—Ì› ‘Õ‰]"
'aFields(5) = "LADING_BILL.CHARGE2 AS [„’«—Ì› »Ê·Ì’…]"
'aFields(6) = "LADING_BILL.DOC_NO AS [—ﬁ„ «·›« Ê—…]"
'aFields(7) = "FORMAT(FILE6_20h.DATE,'yyyy/M/d') AS [ «—ÌŒ «·›« Ê—…]"
'aFields(8) = "FILE6_20H.NAME AS [«”„ «·⁄„Ì·]"
'aFields(9) = "FILE6_20H.PHONE AS [ ·Ì›Ê‰ «·⁄„Ì·]"
'aFields(10) = "FILE6_20H.ADDRESS AS [«·⁄‰Ê«‰]"
'aFields(11) = "LADING_BILL.DOC_NO_RET AS [›« Ê—… «·„— Ã⁄]"
'aFields(12) = "FORMAT(F.DATE,'yyyy/M/d') AS [ «—ÌŒ «·„— Ã⁄]"
'aFields(13) = "FORMAT(LADING_BILL.DATE_DELIVERY,'yyyy/M/d') AS [ «—ÌŒ «·«” ·«„]"
'aFields(14) = "LADING_BILL.ID"

With grid1
    If (cBranch <> "00") Then
        grid1.Editable = flexEDNone
    ElseIf grid1.col = 12 Then
        grid1.Editable = flexEDKbdMouse
    ElseIf grid1.TextMatrix(grid1.Row, 12) <> "" Then
        grid1.Editable = flexEDNone
    ElseIf grid1.TextMatrix(grid1.Row, 14) <> "" Then
        grid1.Editable = flexEDNone
    ElseIf grid1.col = 0 Or grid1.col = 1 Or grid1.col = 2 Or grid1.col = 3 Or grid1.col = 4 Or grid1.col = 5 Then
        grid1.Editable = flexEDKbdMouse
    ElseIf grid1.col = 6 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
        grid1.Editable = flexEDKbdMouse
    ElseIf grid1.col = 7 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
        grid1.Editable = flexEDKbdMouse
    Else
        grid1.Editable = flexEDNone
    End If
End With
End Sub
Private Sub Form_Load()
bEdit = True

openCon con

Set DATA2.Recordset = mycmd(YearsString(Year(Date)), con)
Set xYear.RowSource = DATA2
xYear.ListField = "CODE"
xYear.BoundColumn = "CODE"

Set DATA3.Recordset = mycmd(MonthString(Month(Date)), con)
Set xMonth.RowSource = DATA3
xMonth.ListField = "CODE"
xMonth.BoundColumn = "CODE"

'Set rsShip = myCmd("SELECT *.SHIP FROM SHIP")

Set grid1.DataSource = data11

With grid1
cList1 = StrList2("select code,desca from SHIP WHERE STOPED = 0 order by desca")
cList2 = StrList2("select code,desca from lading_codes order by desca")
End With

bIgClick = True
'LoadText Me
myload
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
bIgClick = False
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
'If cBranch <> "00" Then Exit Sub
If KeyCode = 13 Then
    If col = 0 And grid1.TextMatrix(Row, col) = "" Then Exit Sub
    If col = 3 And grid1.TextMatrix(Row, col) = "" Then Exit Sub
    CellPos KeyCode, Row, col
End If
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
'If cBranch <> "00" Then Exit Sub
If KeyAscii = 13 Then
    If grid1.col = 0 And grid1.TextMatrix(grid1.Row, grid1.col) = "" Then Exit Sub
    If grid1.col = 3 And grid1.TextMatrix(grid1.Row, grid1.col) = "" Then Exit Sub
    KeyAscii = 0
End If
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
If col = 1 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "—ﬁ„ «·»Ê·Ì’… €Ì— „”Ã·"
        Cancel = True
    End If
ElseIf col = 2 Then
    grid1.EditText = myFormat_p(grid1.EditText)
    If (Not IsDate(grid1.EditText)) Then
        MsgBox " «—ÌŒ «·»Ê·Ì’… €Ì— „”Ã·"
        Cancel = True
    End If
ElseIf col = 3 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "‘—ﬂ…«·‘Õ‰ €Ì— „”Ã·…"
        Cancel = True
    End If
ElseIf col = 0 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "‰Ê⁄ «·»Ê·Ì’… €Ì— „”Ã·"
        Cancel = True
    End If
ElseIf col = 4 Or col = 5 Then
    If Not IsNumeric(grid1.EditText) And Trim(grid1.EditText) <> "" Then
        Cancel = True
    End If
ElseIf col = 6 Then
    If Trim(grid1.EditText) <> "" Then
        Dim sReturn As String
        sReturn = myField("select doc_no from file6_20h where online_doc = " & MyParn(grid1.EditText), con) & ""
        If sReturn = "" Then
            MsgBox "·«  ÊÃœ ÿ·»Ì… »Â–« «·—ﬁ„"
            Cancel = True
            Exit Sub
        End If
        If Not invDesc(sReturn, Row, col) Then
            Cancel = True
        Else
            grid1.TextMatrix(Row, 7) = sReturn
        End If
    End If
ElseIf col = 7 Then
    If IsNumeric(grid1.EditText) Then
        If Not invDesc(grid1.EditText, Row, col) Then
            Cancel = True
        End If
    ElseIf Trim(grid1.EditText) = "" Then
        grid1.TextMatrix(Row, 5) = ""
        grid1.TextMatrix(Row, 10) = ""
    Else
        Cancel = True
    End If
ElseIf col = 12 Then
    grid1.EditText = myFormat_p(grid1.EditText)
    If (Not IsDate(grid1.EditText)) And Trim(grid1.EditText) <> "" Then
        Cancel = True
    End If
End If
End Sub
Private Function invDesc(pInv_no As String, Row As Long, col As Long) As Boolean
Dim cmd As New ADODB.command
Set cmd = myCommand("dbo.sp_lading_invoice", con, adStoredProc, AddFlag(Empty, "doc_no", pInv_no))

If IsNull(cmd.Parameters("@date").Value) Then
    MsgBox "—ﬁ„ ›« Ê—… €Ì— ’ÕÌÕ"
    Exit Function
ElseIf cmd.Parameters("@online").Value <> 2 Then
    MsgBox "«·›« Ê—… ·Ì”  «Ê‰ ·«Ì‰"
    'Exit Function
ElseIf cmd.Parameters("@total_item").Value < 0 Then
    MsgBox "›« Ê—… „— Ã⁄"
    'Exit Function
ElseIf Not IsNull(cmd.Parameters("@lading").Value) Then
    If MsgBox("«·›« Ê—… „”Ã· ·ÌÂ« »Ê·Ì’… „”»ﬁ…„‰ ﬁ»·", vbOKCancel + vbDefaultButton2) <> vbOK Then
        Exit Function
    End If
ElseIf IsNull(cmd.Parameters("@SHIP_NO").Value) Then
    MsgBox "«·›« Ê—… ·Ì” ·Â« »Ê·Ì’… ‘Õ‰"
    Exit Function
ElseIf IsNull(cmd.Parameters("@SHIP").Value) Then
    MsgBox "«·›« Ê—… ·Ì” ·Â« ‘—ﬂ… ‘Õ‰"
    Exit Function
ElseIf Not cmd.Parameters("@ISCLOSED").Value Then
    MsgBox "«·›« Ê—… ·Ì”  „€·ﬁ…"
    Exit Function
End If
    
    If col = 7 Then
        grid1.TextMatrix(Row, 6) = cmd.Parameters("@ONLINE_DOC").Value & ""
    Else
        grid1.TextMatrix(Row, 7) = cmd.Parameters("@ONLINE_DOC").Value & ""
    End If
    
    If grid1.TextMatrix(Row, 0) = "0" Then
        grid1.TextMatrix(Row, 1) = cmd.Parameters("@SHIP_NO").Value & ""
        grid1.TextMatrix(Row, 2) = myFormat_p(cmd.Parameters("@DATE").Value)
        grid1.TextMatrix(Row, 3) = cmd.Parameters("@SHIP").Value & ""
        grid1.TextMatrix(Row, 4) = Val(cmd.Parameters("@CHARGE1").Value & "")
        grid1.TextMatrix(Row, 5) = Val(cmd.Parameters("@CHARGE1").Value & "")
    End If
    
    grid1.TextMatrix(Row, 8) = myFormat_p(cmd.Parameters("@DATE").Value)
    grid1.TextMatrix(Row, 9) = cmd.Parameters("@name").Value & ""
    grid1.TextMatrix(Row, 10) = cmd.Parameters("@PHONE").Value & ""
    invDesc = True
End Function
Private Sub fixGrd()
With grid1

'aFields(0) = "LADING_BILL.TYPE AS [‰Ê⁄ «·»Ê·Ì’…]"
'aFields(1) = "LADING_BILL.SHIP_NO AS [—ﬁ„ «·»Ê·Ì’…]"
'aFields(2) = "FORMAT(LADING_BILL.DATE,'yyyy/M/d') AS [ «—ÌŒ «·»Ê·Ì’…]"
'aFields(3) = "LADING_BILL.SHIP AS [‘—ﬂ… «·‘Õ‰]"
'aFields(4) = "LADING_BILL.CHARGE1 AS [„’«—Ì› ‘Õ‰]"
'aFields(5) = "LADING_BILL.CHARGE2 AS [„’«—Ì› »Ê·Ì’…]"
'aFields(6) = "FILE6_20H.ONLINE_DOC AS [—ﬁ„ «·ÿ·»Ì…]"
'aFields(7) = "LADING_BILL.DOC_NO AS [—ﬁ„ «·›« Ê—…]"
'aFields(8) = "FORMAT(FILE6_20h.DATE,'yyyy/M/d') AS [ «—ÌŒ «·›« Ê—…]"
'aFields(9) = "FILE6_20H.NAME AS [«”„ «·⁄„Ì·]"
'aFields(10) = "FILE6_20H.PHONE AS [ ·Ì›Ê‰ «·⁄„Ì·]"
'aFields(11) = "FILE6_20H.ADDRESS AS [«·⁄‰Ê«‰]"
'aFields(12) = "LADING_BILL.DOC_NO_RET AS [›« Ê—… «·„— Ã⁄]"
'aFields(13) = "FORMAT(F.DATE,'yyyy/M/d') AS [ «—ÌŒ «·„— Ã⁄]"
'aFields(14) = "FORMAT(LADING_BILL.DATE_DELIVERY,'yyyy/M/d') AS [ «—ÌŒ «·«” ·«„]"
'aFields(15) = "LADING_BILL.ID"
'
.RowHeight(0) = 600
.ColWidth(0) = 1500
.ColWidth(1) = 1300
.ColWidth(2) = 1300
.ColWidth(3) = 1500
.ColWidth(4) = 1000
.ColWidth(5) = 1000
.ColWidth(7) = 1700
.ColWidth(8) = 1300
.ColWidth(9) = 1500
.ColWidth(10) = 1500
.ColWidth(11) = 5000
.ColWidth(12) = 1600
.ColWidth(13) = 1300

.ColComboList(0) = cList2
.ColComboList(3) = cList1

'.ColHidden(9) = True
.ColHidden(10) = True
.ColHidden(11) = True
.ColHidden(.Cols - 1) = True

For i = 0 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.WordWrap = True

.Cell(flexcpBackColor, 1, 8, .Rows - 1, 10) = &HE0E0E0
.Cell(flexcpBackColor, 1, 13, .Rows - 1, 15) = &HE0E0E0
End With
End Sub
Private Sub CalcTotals()
panel1(0).Caption = IIf(grid1.Rows > 2, "⁄œœ «·«Ì’«·«  " & grid1.Rows - 2, "·« ÌÊÃœ «Ì’«·« ")
End Sub
Private Function validRow(Row As Long) As Boolean
With grid1
If Trim(grid1.TextMatrix(Row, 0)) = "" Then Exit Function
If Not IsNumeric(grid1.TextMatrix(Row, 1)) Then Exit Function
If Not IsDate(grid1.TextMatrix(Row, 2)) Then Exit Function
If Not IsNumeric(grid1.TextMatrix(Row, 3)) Then Exit Function

If grid1.TextMatrix(Row, 7) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
'If cBranch <> "00" Then Exit Sub
If KeyCode = 13 Then
     CellPos KeyCode, grid1.Row, grid1.col
ElseIf KeyCode = 112 And grid1.col = 7 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
    SalesLadingLookup Me, oSearchInv, , , IIf(grid1.TextMatrix(grid1.Row, 7) <> "", "»œÊ‰ ›« Ê—…", "")
ElseIf KeyCode = 46 And grid1.Row <> 0 And grid1.Row <> grid1.Rows - 1 Then
    If Trim(grid1.TextMatrix(grid1.Row, grid1.Cols - 1)) <> "" Then
        If MsgBox("Õ–› «·«Ì’«· ø", vbOKCancel + vbDefaultButton2) = vbOK Then
            If grid1.TextMatrix(grid1.Row, 12) <> "" Then
                MsgBox "··»Ê·Ì’… „” ‰œ „— Ã⁄"
                Exit Sub
            ElseIf grid1.TextMatrix(grid1.Row, 14) <> "" Then
                MsgBox "··»Ê·’…  «—ÌŒ «” ·«„"
                Exit Sub
            End If
            
            On Error GoTo myerror
            Dim nAffect As Long
            con.Execute "delete from LADING_BILL " & _
                        " where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1) & _
                        " AND DOC_NO_RET IS NULL" & _
                        " AND DATE_DELIVERY IS NULL", nAffect
            If nAffect = 0 Then
                MsgBox "·„ Ì „ Õ–› «·»Ê·Ì’… !! ﬁœ ÌﬂÊ‰ ·Â« ›« Ê—… „— Ã⁄ «Ê  «—ÌŒ «” ·«„"
                Exit Sub
            End If
            
            grid1.RemoveItem grid1.Row
            grid1.SetFocus
            CalcTotals
            bStopCell = True
        End If
    End If
ElseIf KeyCode = 112 And grid1.col = 3 Then
    'SupLookupAll Me, oSearchSup
End If
Exit Sub
myerror:
If Err.Number <> 0 Then MsgBox Err.Description
myLoadGrd
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal col As Long)
KeyCode = 0
If bStopCell Then
    bStopCell = False
ElseIf col < 7 Then
    If grid1.TextMatrix(Row, 0) = "0" Then
        grid1.Select Row, 6
        grid1.ShowCell grid1.Row, grid1.col
    Else
        grid1.Select Row, col + 1
        grid1.ShowCell grid1.Row, grid1.col
    End If
ElseIf col < 12 And grid1.TextMatrix(Row, 12) = "" Then
    grid1.Select Row, 12
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 0, 12)
    grid1.ShowCell grid1.Row, grid1.FrozenCols
Else
    grid1.Select Row, col
End If
End Sub
Private Sub myAddItem()
With grid1
    .AddItem ""
    If cmdType.Tag <> "" Then
        grid1.TextMatrix(grid1.Rows - 1, 3) = cmdType.Tag
'    ElseIf grid1.Rows > 2 Then
'        grid1.TextMatrix(grid1.Rows - 1, 3) = grid1.TextMatrix(grid1.Rows - 2, 3)
    End If
End With
End Sub
Private Function myreplace(Row As Long, ByRef nId As Variant) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(aInsert, "[TYPE]", addvalue(grid1.TextMatrix(Row, 0)))
aInsert = AddFlag(aInsert, "[SHIP_NO]", addstring(grid1.TextMatrix(Row, 1)))
aInsert = AddFlag(aInsert, "[DATE]", addDate(grid1.TextMatrix(Row, 2)))
aInsert = AddFlag(aInsert, "[SHIP]", addstring(grid1.TextMatrix(Row, 3)))
aInsert = AddFlag(aInsert, "[CHARGE1]", grid1.ValueMatrix(Row, 4))
aInsert = AddFlag(aInsert, "[CHARGE2]", grid1.ValueMatrix(Row, 5))
aInsert = AddFlag(aInsert, "[doc_no]", addstring(grid1.TextMatrix(Row, 7)))
aInsert = AddFlag(aInsert, "[DATE_MAIL]", addDate(grid1.TextMatrix(Row, 12)))
con.BeginTrans

On Error GoTo myerror
If grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
    aInsert = AddFlag(aInsert, "[HOST_NAME]", addstring(GetComputerName))
    con.Execute addInsert(aInsert, "LADING_BILL")
    nId = myField("SELECT SCOPE_IDENTITY()", con)
Else
    con.Execute addUpdate(aInsert, "LADING_BILL", "ID = " & grid1.TextMatrix(Row, grid1.Cols - 1))
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Function
Public Sub myProc()
If ActiveControl.Name = grid1.Name And grid1.col = 7 Then
    If Not invDesc(oSearchInv.grid1.TextMatrix(oSearchInv.grid1.Row, 0), grid1.Row, grid1.col) Then Exit Sub
    grid1.TextMatrix(grid1.Row, 7) = oSearchInv.grid1.TextMatrix(oSearchInv.grid1.Row, 0)
    Grid1_AfterEdit grid1.Row, 7
    CellPos 13, grid1.Row, 7
    oSearchInv.Hide
ElseIf ActiveControl.Name = cmdShip.Name Then
    If oSearchship.grid1.TextMatrix(oSearchship.grid1.Row, 0) = "" Then
        cmdShip.Tag = ""
        cmdShip.Caption = cmdShip.TagVariant
    Else
        cmdShip.Tag = oSearchship.grid1.TextMatrix(oSearchship.grid1.Row, 0)
        cmdShip.Caption = oSearchship.grid1.TextMatrix(oSearchship.grid1.Row, 1)
    End If
    myload
    oSearchship.Hide
ElseIf ActiveControl.Name = cmdType.Name Then
    If oSearchLading.grid1.TextMatrix(oSearchLading.grid1.Row, 0) = "" Then
        cmdType.Tag = ""
        cmdType.Caption = cmdType.TagVariant
    Else
        cmdType.Tag = oSearchLading.grid1.TextMatrix(oSearchLading.grid1.Row, 0)
        cmdType.Caption = oSearchLading.grid1.TextMatrix(oSearchLading.grid1.Row, 1)
    End If
    myload
    oSearchLading.Hide
End If
End Sub
Private Sub optclose_Click(Index As Integer)
If bIgClick Then Exit Sub
myload
End Sub

Private Sub optmail_Click(Index As Integer)
If bIgClick Then Exit Sub
myLoadGrd
End Sub

Private Sub xDoc_no_Change()
'myLoadGrd
End Sub
Private Function retWhere() As String
Dim cWhere As New ChilkatStringBuilder
If Not (IsDate(xdate1.text) Or IsDate(xdate2.text)) Then
    If xMonth.BoundText <> "" And xYear.BoundText <> "" Then
        cWhere.Append " AND LADING_BILL.YEAR_MONTH = " & MyParn(xYear.BoundText & "-" & RetZero(xMonth.BoundText, 2))
    ElseIf xYear.BoundText <> "" Then
        cWhere.Append " AND LADING_BILL.YEAR_MONTH LIKE " & MyParn(xYear.BoundText & "%")
    End If
End If

If IsDate(xdate1.text) Then
    cWhere.Append " AND " & "LADING_BILL.DATE >= " & addDate(xdate1.text)
End If

If IsDate(xdate2.text) Then
    cWhere.Append " AND " & "LADING_BILL.DATE <= " & addDate(xdate2.text)
End If

If chkDay.Value = 1 Then
    cWhere.Append " AND " & "LADING_BILL.DATE = " & addDate(Date)
End If

If optClose(1).Value Then
    cWhere.Append " AND (LADING_BILL.CLOSED = 0)"
ElseIf optClose(2).Value Then
    cWhere.Append " AND (LADING_BILL.CLOSED = 1)"
End If

If Trim(xdoc_no.text) <> "" Then
    cWhere.Append " AND " & "LADING_BILL.DOC_NO = " & MyParn(xdoc_no.text)
End If

If Trim(xdoc_no_ret.text) <> "" Then
    cWhere.Append " AND " & "LADING_BILL.DOC_NO_RET = " & MyParn(xdoc_no.text)
End If

If Trim(xOnline_doc.text) <> "" Then
    cWhere.Append " AND " & "FILE6_20H.ONLINE_DOC = " & MyParn(xOnline_doc.text)
End If

If Trim(xship_no.text) <> "" Then
    cWhere.Append " AND " & MyParnAnd(xship_no.text, "LADING_BILL.SHIP_NO")
End If

If optMail(1).Value Then
    cWhere.Append " AND (DATE_MAIL IS NOT NULL)"
ElseIf optMail(2).Value Then
    cWhere.Append " AND (DATE_MAIL IS NULL)"
End If

If IsDate(xdate_mail1.text) Then
    cWhere.Append " AND " & "LADING_BILL.DATE_MAIL >= " & addDate(xdate_mail1.text)
End If

If IsDate(xdate_mail2.text) Then
    cWhere.Append " AND " & "LADING_BILL.DATE_MAIL <= " & addDate(xdate_mail2.text)
End If

If chkdayMail.Value = 1 Then
    cWhere.Append " AND " & "LADING_BILL.DATE_MAIL = " & addDate(Date)
End If

If IsDate(xdate_Delivery1.text) Then
    cWhere.Append " AND " & "LADING_BILL.date_delivery >= " & addDate(xdate_Delivery1.text)
End If

If IsDate(xdate_Delivery2.text) Then
    cWhere.Append " AND " & "LADING_BILL.date_delivery <= " & addDate(xdate_Delivery2.text)
End If
'
If chkDay_delivery.Value = 1 Then
    cWhere.Append " AND " & "LADING_BILL.DATE_DELIVERY = " & addDate(Date)
End If

If IsDate(xdate_return1.text) Then
    cWhere.Append " AND " & "F.DATE >= " & addDate(xdate_return1.text)
End If

If IsDate(xdate_Return2.text) Then
    cWhere.Append " AND " & "F.DATE <= " & addDate(xdate_Return2.text)
End If
'
If chkDay_delivery.Value = 1 Then
    cWhere.Append " AND " & "F.DATE = " & addDate(Date)
End If

If Trim(xName.text) <> "" Then
    cWhere.Append " AND " & MyParnAnd(xName.text, "FILE6_20H.NAME")
End If

If Trim(xPhone.text) <> "" Then
    cWhere.Append " AND " & MyParnAnd(xPhone.text, "FILE6_20H.PHONE")
End If

If cmdType.Tag <> "" Then
    cWhere.Append " AND " & "LADING_BILL.TYPE = " & cmdType.Tag
End If

If cmdShip.Tag <> "" Then
    cWhere.Append " AND " & "LADING_BILL.SHIP = " & MyParn(cmdShip.Tag)
End If
retWhere = Mid(cWhere.GetAsString, 6)
End Function
Private Sub xYear_Change()
If bIgClick Then Exit Sub
myLoadGrd
End Sub
Private Sub xMonth_Change()
If bIgClick Then Exit Sub
If xYear.text = "" Then Exit Sub
myLoadGrd
End Sub
Private Sub xYear_GotFocus()
myGotFocus xYear
End Sub
Private Sub xYear_LostFocus()
myLostFocus xYear
If Not xYear.MatchedWithList Then xYear.BoundText = ""
End Sub
Private Sub xMonth_GotFocus()
myGotFocus xMonth
End Sub
Private Sub xMonth_LostFocus()
myLostFocus xMonth
If Not xMonth.MatchedWithList Then xMonth.BoundText = ""
End Sub
Private Sub xdate1_GotFocus()
myGotFocus xdate1
End Sub
Private Sub xdate1_LostFocus()
myLostFocus xdate1
myValidDate xdate1
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xdate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xdate2
myValidDate xdate2
End Sub
Private Sub XONLINE_DOC_GotFocus()
myGotFocus xOnline_doc
End Sub
Private Sub XONLINE_DOC_LostFocus()
myLostFocus xOnline_doc
End Sub
Private Sub xdoc_no_ret_GotFocus()
myGotFocus xdoc_no_ret
End Sub
Private Sub xdoc_no_ret_LostFocus()
myLostFocus xdoc_no_ret
End Sub
Private Sub xship_no_GotFocus()
myGotFocus xship_no
End Sub
Private Sub xship_no_LostFocus()
myLostFocus xship_no
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xdoc_no
End Sub
Private Sub xDoc_No_LostFocus()
myLostFocus xdoc_no
End Sub
Private Sub xName_GotFocus()
myGotFocus xName
End Sub
Private Sub xname_LostFocus()
myLostFocus xName
End Sub
Private Sub xPhone_GotFocus()
myGotFocus xPhone
End Sub
Private Sub XPHONE_LostFocus()
myLostFocus xPhone
End Sub
Private Sub xdate_return1_GotFocus()
myGotFocus xdate_return1
End Sub
Private Sub xdate_return1_LostFocus()
myLostFocus xdate_return1
myValidDate xdate_return1
End Sub
Private Sub xdate_Return2_GotFocus()
myGotFocus xdate_Return2
End Sub
Private Sub xdate_Return2_LostFocus()
myLostFocus xdate_Return2
myValidDate xdate_Return2
End Sub
Private Sub xdate_Delivery2_GotFocus()
myGotFocus xdate_Delivery2
End Sub
Private Sub xdate_Delivery2_LostFocus()
myLostFocus xdate_Delivery2
myValidDate xdate_Delivery2
End Sub
Private Sub xdate_Delivery1_GotFocus()
myGotFocus xdate_Delivery1
End Sub
Private Sub xdate_Delivery1_LostFocus()
myLostFocus xdate_Delivery1
myValidDate xdate_Delivery1
End Sub
Private Sub xdate_Mail1_GotFocus()
myGotFocus xdate_mail1
End Sub
Private Sub xdate_Mail1_LostFocus()
myLostFocus xdate_mail1
myValidDate xdate_mail1
End Sub
Private Sub xdate_Mail2_GotFocus()
myGotFocus xdate_mail2
End Sub
Private Sub xdate_Mail2_LostFocus()
myLostFocus xdate_mail2
myValidDate xdate_mail2
End Sub

