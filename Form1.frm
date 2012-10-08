VERSION 5.00
Object = "{962F707C-C66B-11D6-AFC5-0050BACCDC45}#1.0#0"; "PISO813X.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00808000&
   Caption         =   "AD Demo , Ads_Float()"
   ClientHeight    =   7665
   ClientLeft      =   1530
   ClientTop       =   2160
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   8355
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF8080&
      Caption         =   "Channel"
      Height          =   2295
      Left            =   120
      TabIndex        =   26
      Top             =   5160
      Width           =   6135
      Begin VB.CheckBox Check1 
         Caption         =   "0"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   59
         Top             =   360
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "1"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   58
         Top             =   360
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "2"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   57
         Top             =   360
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "3"
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   56
         Top             =   360
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "4"
         Height          =   255
         Index           =   4
         Left            =   3240
         TabIndex        =   55
         Top             =   360
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "5"
         Height          =   255
         Index           =   5
         Left            =   3960
         TabIndex        =   54
         Top             =   360
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "6"
         Height          =   255
         Index           =   6
         Left            =   4680
         TabIndex        =   53
         Top             =   360
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "7"
         Height          =   255
         Index           =   7
         Left            =   5400
         TabIndex        =   52
         Top             =   360
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "8"
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   51
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "9"
         Height          =   255
         Index           =   9
         Left            =   1080
         TabIndex        =   50
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "10"
         Height          =   255
         Index           =   10
         Left            =   1800
         TabIndex        =   49
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "11"
         Height          =   255
         Index           =   11
         Left            =   2520
         TabIndex        =   48
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "12"
         Height          =   255
         Index           =   12
         Left            =   3240
         TabIndex        =   47
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "19"
         Height          =   255
         Index           =   19
         Left            =   2520
         TabIndex        =   46
         Top             =   1080
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "18"
         Height          =   255
         Index           =   18
         Left            =   1800
         TabIndex        =   45
         Top             =   1080
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "17"
         Height          =   255
         Index           =   17
         Left            =   1080
         TabIndex        =   44
         Top             =   1080
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "31"
         Height          =   255
         Index           =   31
         Left            =   5400
         TabIndex        =   43
         Top             =   1440
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "30"
         Height          =   255
         Index           =   30
         Left            =   4680
         TabIndex        =   42
         Top             =   1440
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "29"
         Height          =   255
         Index           =   29
         Left            =   3960
         TabIndex        =   41
         Top             =   1440
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "28"
         Height          =   255
         Index           =   28
         Left            =   3240
         TabIndex        =   40
         Top             =   1440
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "27"
         Height          =   255
         Index           =   27
         Left            =   2520
         TabIndex        =   39
         Top             =   1440
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "26"
         Height          =   255
         Index           =   26
         Left            =   1800
         TabIndex        =   38
         Top             =   1440
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "25"
         Height          =   255
         Index           =   25
         Left            =   1080
         TabIndex        =   37
         Top             =   1440
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "24"
         Height          =   255
         Index           =   24
         Left            =   360
         TabIndex        =   36
         Top             =   1440
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "23"
         Height          =   255
         Index           =   23
         Left            =   5400
         TabIndex        =   35
         Top             =   1080
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "22"
         Height          =   255
         Index           =   22
         Left            =   4680
         TabIndex        =   34
         Top             =   1080
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "21"
         Height          =   255
         Index           =   21
         Left            =   3960
         TabIndex        =   33
         Top             =   1080
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "15"
         Height          =   255
         Index           =   15
         Left            =   5400
         TabIndex        =   32
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "14"
         Height          =   255
         Index           =   14
         Left            =   4680
         TabIndex        =   31
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "13"
         Height          =   255
         Index           =   13
         Left            =   3960
         TabIndex        =   30
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "20"
         Height          =   255
         Index           =   20
         Left            =   3240
         TabIndex        =   29
         Top             =   1080
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "16"
         Height          =   255
         Index           =   16
         Left            =   360
         TabIndex        =   28
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox tbChannelItv 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   288
         Left            =   1920
         TabIndex        =   27
         Text            =   "10"
         Top             =   1840
         Width           =   1092
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF8080&
         Caption         =   "Channel Interval:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   60
         Top             =   1920
         Width           =   1575
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7200
      Top             =   4320
   End
   Begin PISO813XLib.PISO813X PISO813X1 
      Height          =   375
      Left            =   480
      TabIndex        =   21
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   0
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Caption         =   "Misc Settings"
      Height          =   1935
      Left            =   3240
      TabIndex        =   16
      Top             =   3120
      Width           =   3012
      Begin VB.TextBox ItvTime 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   288
         Left            =   1440
         TabIndex        =   25
         Text            =   "180"
         Top             =   1040
         Width           =   1092
      End
      Begin VB.TextBox dataFile 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   288
         Left            =   1440
         TabIndex        =   23
         Text            =   "D:\my.xls"
         Top             =   1440
         Width           =   1092
      End
      Begin VB.TextBox eAD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   288
         Left            =   1440
         TabIndex        =   18
         Top             =   680
         Width           =   1092
      End
      Begin VB.ComboBox InRange 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "Form1.frx":0000
         Left            =   1440
         List            =   "Form1.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   240
         Width           =   1308
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF8080&
         Caption         =   "Interval"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Save Data:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   870
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF8080&
         Caption         =   "AD (float)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Input Range"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   228
         Index           =   3
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1056
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0FF&
      Caption         =   " Hardware Setting "
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   3012
      Begin VB.VScrollBar VScroll1 
         Height          =   372
         Left            =   2520
         TabIndex        =   14
         Top             =   600
         Width           =   252
      End
      Begin VB.TextBox eSelectBoard 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         Height          =   372
         Left            =   1572
         TabIndex        =   13
         Text            =   "0"
         Top             =   588
         Width           =   972
      End
      Begin VB.TextBox eTotalBoards 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   288
         Left            =   1560
         TabIndex        =   11
         Text            =   "0"
         Top             =   240
         Width           =   972
      End
      Begin VB.ComboBox cbBiUni 
         BackColor       =   &H00C0FFC0&
         Height          =   288
         ItemData        =   "Form1.frx":0004
         Left            =   1560
         List            =   "Form1.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1440
         Width           =   1212
      End
      Begin VB.ComboBox cbJmp20v 
         BackColor       =   &H00C0FFC0&
         Height          =   288
         ItemData        =   "Form1.frx":002B
         Left            =   1560
         List            =   "Form1.frx":0035
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   1212
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Choose a Board Number to Active"
         ForeColor       =   &H00FF0000&
         Height          =   372
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1332
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Total Boards : "
         ForeColor       =   &H00FF0000&
         Height          =   252
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   1092
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "JP2 Setting "
         ForeColor       =   &H00FF0000&
         Height          =   252
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   1092
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0FF&
         Caption         =   "JP1 Setting"
         ForeColor       =   &H00FF0000&
         Height          =   252
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1212
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Analog to Digital Display"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2892
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8028
      Begin VB.TextBox YScale 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   270
         TabIndex        =   5
         Text            =   "-5"
         Top             =   2280
         Width           =   552
      End
      Begin VB.TextBox YScale 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   270
         TabIndex        =   4
         Text            =   "5"
         Top             =   480
         Width           =   552
      End
      Begin VB.PictureBox Gph 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   2148
         Left            =   960
         ScaleHeight     =   2085
         ScaleWidth      =   6855
         TabIndex        =   3
         Top             =   480
         Width           =   6912
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6600
      Top             =   4320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Active"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6600
      TabIndex        =   1
      Top             =   3360
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6600
      TabIndex        =   0
      Top             =   3840
      Width           =   1332
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xlsApp As Excel.Application
Dim xlsWorkbook As Excel.Workbook
Dim xlsSheet As Excel.Worksheet

Dim TotalTimeInterval As Long
Dim TotalTimePassed As Long
Dim ChannelTimeInterval As Long
Dim ChannelTimePassed As Long
Dim TotalChannelChkdNo As Integer
Dim CurrentChannelIndex As Integer

Dim iColIndex As Integer

Dim wInitialCode As Integer
Dim YS(0 To 1) As Single

Dim wCfgCode As Integer
Dim wBiUni, wJmp20v As Integer
Dim dwCount As Long
Dim bProcessing As Boolean

Dim wBaseAddr                    As Long
Dim wIrqNo                       As Integer
Dim wSubVendor, wSubDevice       As Integer
Dim wSubAux, wSlotBus            As Integer
Dim wSlotDevice                  As Integer
Dim wSelectBoard, wTotalBoards   As Integer
Dim wRetVal                      As Integer

Private Sub cbBiUni_Click()
    SetInputRange
End Sub

Private Sub cbJmp20v_Click()
    If cbJmp20v.ListIndex = 1 Then
       cbBiUni.ListIndex = 1      '// Not Support 20v + Unipolar
       cbBiUni.Enabled = False
    Else
       cbBiUni.Enabled = True
    End If
    
    SetInputRange
End Sub


Private Sub InRange_Click()
  Dim G, G2 As Single
  
  wCfgCode = InRange.ListIndex  'While Input Range is changed, reset the Gain Code
  
  If (cbJmp20v.ListIndex = 0) Then     ' 10v
    If (cbBiUni.ListIndex = 0) Then    'Unipolar
        G = 10# / 2 ^ InRange.ListIndex
    Else
        G = 5# / 2 ^ InRange.ListIndex
    End If
  Else                                ' 20v
    If (cbBiUni.ListIndex = 0) Then   'Unipolar
        G = 20# / 2 ^ InRange.ListIndex
        If (InRange.ListIndex = 0) Then
            G = 0#
        End If
    Else
        G = 10# / 2 ^ InRange.ListIndex
    End If
  End If

  If (cbBiUni.ListIndex = 0) Then    ' Unipolar
    G2 = 0#
  Else
    G2 = G * -1
  End If
  
  
  YScale(0).Text = G
  YScale(1).Text = G2
  YScale_LostFocus (0)
  YScale_LostFocus (1)
End Sub

Private Sub PISO813X1_OnError(ByVal lErrorCode As Long)
MsgBox "Error Code: " + Str(lErrorCode) + Chr(13) _
          + "Error Message: " + PISO813X1.ErrorString
End Sub

Private Sub Timer2_Timer()
    Dim i, j As Integer

    If bProcessing = True Then
        Exit Sub
    Else
        bProcessing = True
    End If
    
        If TotalTimePassed < TotalTimeInterval Then
        TotalTimePassed = TotalTimePassed + Timer2.Interval
    Else
        TotalTimePassed = Timer2.Interval
        CurrentChannelIndex = 1
        ChannelTimePassed = ChannelTimeInterval
    End If
    
    
    If CurrentChannelIndex <= TotalChannelChkdNo Then
        If ChannelTimePassed < ChannelTimeInterval Then
            ChannelTimePassed = ChannelTimePassed + Timer2.Interval
        Else
            ChannelTimePassed = Timer2.Interval
            
            j = 0
            For i = 0 To 31
                If Check1(i).Value = 1 Then
                    j = j + 1
                End If
                If j = CurrentChannelIndex Then
                    Exit For
                End If
            Next i
            
            RecordChannel i
            CurrentChannelIndex = CurrentChannelIndex + 1
        End If
    End If
    
    bProcessing = False
End Sub

Private Sub VScroll1_Change()
    eSelectBoard.Text = VScroll1.Value
End Sub

Private Sub YScale_LostFocus(Index As Integer)
  YScale(Index).Text = Trim(Val(YScale(Index).Text))
  If Val(YScale(Index).Text) > 12 Then YScale(Index).Text = "12"
  If Val(YScale(Index).Text) < -12 Then YScale(Index).Text = "-12"
  YS(Index) = Val(YScale(Index).Text)
  If YS(Index) = YS((Index + 1) Mod 2) Then
    YS(Index) = YS(Index) + IIf(Index = 0, 0.1, -0.1)
  End If
  YScale(Index).Text = Trim(YS(Index))
  Gph.Cls
  Gph.Scale (0, YS(0))-(100, YS(1))
End Sub


Private Sub Command1_Click()
    Timer1.Enabled = False
    PISO813X1.DriverClose
    End
End Sub


Private Sub Command2_Click()
    Dim wRetVal As Integer
    Dim i As Integer

    If Command2.Caption = "Active" Then
    
        If dataFile.Text = "" Then
            MsgBox "Invalid path to save data recorded."
            Exit Sub
        End If
        
        TotalChannelChkdNo = 0
        ChannelTimeInterval = Val(tbChannelItv.Text) * Timer2.Interval
        TotalTimeInterval = Val(ItvTime.Text) * Timer2.Interval
        TotalTimePassed = TotalTimeInterval
        
        For i = 0 To 31
            Check1(i).Enabled = False
            If Check1(i).Value = 1 Then
                TotalChannelChkdNo = TotalChannelChkdNo + 1
            End If
        Next i
        
        If TotalChannelChkdNo = 0 Then
            MsgBox "Pls choose a channel to detect!"
            Exit Sub
        End If
        
        If (TotalChannelChkdNo * ChannelTimeInterval) > TotalTimeInterval Then
            MsgBox "Channel intervals of all channels together should be less than Interval(Total)!"
            Exit Sub
        End If
        
        wSelectBoard = Val(eSelectBoard.Text)
        If wSelectBoard > Val(eTotalBoards.Text) - 1 _
        Or wSelectBoard < 0 Then
            MsgBox "Invalid board number, Pls retry!!"
            Exit Sub
        End If
        'Get board's Configuration Space
        PISO813X1.ActiveBoard = Val(eSelectBoard.Text)

        '************************************************************
        ' enable all DI/DO port
        '************************************************************
        
        wCfgCode = InRange.ListIndex
        
        Command2.Caption = "Stop"
        Command1.Enabled = False
        
        bProcessing = False
        Timer1.Enabled = True
        Timer2.Enabled = True
    Else
        For i = 0 To 31
            Check1(i).Enabled = True
        Next i
    
        Timer1.Enabled = False
        Timer2.Enabled = False
        Command2.Caption = "Active"
        Command1.Enabled = True
    End If
    
    wBiUni = cbBiUni.ListIndex
    wJmp20v = cbJmp20v.ListIndex
End Sub

Private Sub Form_Load()
    Dim rtn
    
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
   
    Command2.Caption = "Active"
    
    wTotalBoards = PISO813X1.DriverInit
    
        Command2.Enabled = True
        Command1.Enabled = True
    
    eTotalBoards.Text = wTotalBoards
    VScroll1.Min = wTotalBoards - 1
    VScroll1.Max = 0
    
    cbBiUni.ListIndex = 1
    cbJmp20v.ListIndex = 1
    
    YS(0) = 5: YS(1) = -5
    wCfgCode = InRange.ListIndex
    
    SetInputRange
    InRange_Click
    
    TotalTimeInterval = 180000
    TotalTimePassed = 1000 '1800 * 1000
    ChannelTimeInterval = 10000 '10 * 1000
    ChannelTimePassed = 1000 '1000
End Sub

Private Sub InitXls()
    Set xlsApp = CreateObject("Excel.Application")
    Set xlsWorkbook = xlsApp.Workbooks.Open(dataFile.Text)
    Set xlsSheet = xlsWorkbook.Worksheets(1)
    xlsApp.DisplayAlerts = False
    xlsApp.Visible = False
End Sub

Private Sub FinitXls()
    If Not xlsWorkbook Is Nothing Then
        xlsWorkbook.Save
        xlsWorkbook.Close
    End If
    
    If Not xlsApp Is Nothing Then
        xlsApp.Quit
    End If
    
    Set xlsSheet = Nothing
    Set xlsWorkbook = Nothing
    Set xlsApp = Nothing
End Sub

Private Sub RecordIntoXls(fVal As Single, iCNo As Integer)
    Dim rowIndex As Integer

    InitXls
    
    rowIndex = 1
    Do While xlsSheet.Cells(rowIndex, iColIndex + 1) <> ""
        rowIndex = rowIndex + 1
    Loop
    
    If rowIndex = 1 Then
        xlsSheet.Cells(1, iColIndex + 1) = "Channel: " & iCNo
        rowIndex = 2
    End If
    
    xlsSheet.Cells(rowIndex, iColIndex + 1) = Format(Now, "yyyy-mm-dd hh:mm:ss")
    xlsSheet.Cells(rowIndex, iColIndex + 2) = Format(fVal, "###,###.000")
    
    FinitXls
End Sub

Private Sub RecordChannel(ByVal ChannelNo As Integer)
    Dim fVal As Single
    
    PISO813X1.SetChannelGain ChannelNo, wCfgCode
    iColIndex = ChannelNo * 2
    
    fVal = PISO813X1.AnalogIn(wJmp20v, wBiUni)
'    fVal = 12.56
    
    If Dir(dataFile.Text) <> "" Then
        RecordIntoXls fVal, ChannelNo
    Else
        MsgBox "File: " + dataFile.Text + " does not exist!"
    End If
End Sub

Private Sub Timer1_Timer()
    Dim i, iRtn As Integer
    Dim fVal As Single
    Dim fBuf(1000) As Single
    
    If bProcessing = True Then
        Exit Sub
    Else
        bProcessing = True
    End If

    fVal = PISO813X1.AnalogIn(wJmp20v, wBiUni)
    eAD.Text = Format(fVal, "###,###.000")

    PISO813X1.AnalogInMulti wJmp20v, wBiUni, fBuf(0), 200

    Gph.Cls
    Gph.PSet (0, fBuf(0))
    For i = 0 To 199
        Gph.Line -((i - 1), fBuf(i))
        DoEvents
    Next i
    
    bProcessing = False
    
End Sub

Private Sub SetInputRange()

    InRange.Clear

    If (cbJmp20v.ListIndex = 1) Then   ' 20v
        If (cbBiUni.ListIndex = 0) Then   ' Unipolar
            InRange.AddItem " Not Use"    ' Not Support 20v + Unipolar
        Else
            InRange.AddItem "    -10~10   "
            InRange.AddItem "     -5~5    "
            InRange.AddItem "   -2.5~2.5  "
            InRange.AddItem "  -1.25~1.25 "
            InRange.AddItem " -0.625~0.625"
        End If
    Else                             ' 10v
        If (cbBiUni.ListIndex = 0) Then   ' Unipolar
            InRange.AddItem " 0~10   "
            InRange.AddItem " 0~5    "
            InRange.AddItem " 0~2.5  "
            InRange.AddItem " 0~1.25 "
            InRange.AddItem " 0~0.625"
        Else
            InRange.AddItem "     -5~5     "
            InRange.AddItem "   -2.5~2.5   "
            InRange.AddItem "  -1.25~1.25  "
            InRange.AddItem " -0.625~0.625 "
        End If
    End If

    InRange.ListIndex = 0                 'Reset the Input Range
    If (cbJmp20v.ListIndex = 0) Then      ' 20v
        If (cbBiUni.ListIndex = 0) Then   ' Unipolar
            InRange.ListIndex = 1         ' GainCode 0 : Not used
        End If
    End If
    InRange_Click
End Sub


