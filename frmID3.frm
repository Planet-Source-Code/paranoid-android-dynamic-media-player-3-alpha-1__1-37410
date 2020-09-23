VERSION 5.00
Object = "{49616C06-65BE-11D4-804C-000021F0FF9D}#1.0#0"; "ALFAFISHEZYID33VB6.OCX"
Begin VB.Form frmId3 
   BackColor       =   &H00A25100&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ID3 Tag Editor"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   Icon            =   "frmID3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraID3v2 
      BackColor       =   &H008C4600&
      Caption         =   "ID3v2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3555
      Left            =   4500
      TabIndex        =   17
      Top             =   60
      Width           =   4335
      Begin VB.ComboBox cboFrames 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1650
         Width           =   4095
      End
      Begin VB.ListBox lstFrames 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         Left            =   135
         TabIndex        =   22
         Top             =   2040
         Width           =   4095
      End
      Begin VB.CommandButton cmdNextPic 
         Caption         =   "->"
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevPic 
         Caption         =   "<-"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblPicDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   1560
         TabIndex        =   21
         Top             =   240
         Width           =   45
      End
      Begin VB.Label lblPicNr 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1320
         TabIndex        =   20
         Top             =   240
         Width           =   255
      End
      Begin VB.Image imgPicture 
         Height          =   255
         Left            =   120
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.Frame fraInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H008C4600&
      Caption         =   "MPEG Info"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2955
      Left            =   2220
      TabIndex        =   15
      Top             =   60
      Width           =   2175
      Begin AlfafishEzyID33.EzyID3 ID3 
         Left            =   90
         Top             =   225
         _ExtentX        =   1005
         _ExtentY        =   1005
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2595
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.ComboBox cboGenre 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2700
      Width           =   2055
   End
   Begin VB.TextBox txtComment 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      MaxLength       =   30
      TabIndex        =   3
      Top             =   2100
      Width           =   2055
   End
   Begin VB.TextBox txtYear 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      MaxLength       =   30
      TabIndex        =   5
      Top             =   3300
      Width           =   975
   End
   Begin VB.TextBox txtAlbum 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      MaxLength       =   30
      TabIndex        =   2
      Top             =   1500
      Width           =   2055
   End
   Begin VB.TextBox txtSong 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      MaxLength       =   30
      TabIndex        =   0
      Top             =   300
      Width           =   2055
   End
   Begin VB.TextBox txtArtist 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      MaxLength       =   30
      TabIndex        =   1
      Top             =   900
      Width           =   2055
   End
   Begin VB.TextBox txtTrack 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1140
      MaxLength       =   30
      TabIndex        =   6
      Text            =   "0"
      Top             =   3300
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save"
      Height          =   495
      Left            =   2220
      TabIndex        =   7
      Top             =   3105
      Width           =   2175
   End
   Begin VB.Label lblGenre 
      BackStyle       =   0  'Transparent
      Caption         =   "Genre"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   14
      Top             =   2460
      Width           =   1035
   End
   Begin VB.Label lblComment 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   13
      Top             =   1860
      Width           =   1035
   End
   Begin VB.Label lblYear 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   12
      Top             =   3060
      Width           =   435
   End
   Begin VB.Label lblAlbum 
      BackStyle       =   0  'Transparent
      Caption         =   "Album"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   11
      Top             =   1260
      Width           =   1035
   End
   Begin VB.Label lblSong 
      BackStyle       =   0  'Transparent
      Caption         =   "Song"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   1035
   End
   Begin VB.Label lblArtist 
      BackStyle       =   0  'Transparent
      Caption         =   "Artist"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   660
      Width           =   1035
   End
   Begin VB.Label lblTrack 
      BackStyle       =   0  'Transparent
      Caption         =   "Track"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1140
      TabIndex        =   8
      Top             =   3060
      Width           =   465
   End
End
Attribute VB_Name = "frmId3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This code was not written by me, this was
'created by Erik Christiansson, AlfaFish.

'So I dont have any comments throughout it.

'Sorry

Private Sub cboFrames_Click()
Dim f As Integer

If ID3.HaveTag >= 2 Then
    lstFrames.Clear
    
    Select Case cboFrames.Text
        Case "Text"
            For f = 1 To ID3.TextFrameNum
                If ID3.TextFrameName(f) = "TCON" Then _
                    lstFrames.AddItem ID3.TextFrameName(f) & " (decoded): " & ID3.DecodeTCON(ID3.TextFrame(f))
                lstFrames.AddItem ID3.TextFrameName(f) & ": " & ID3.TextFrame(f)
            Next f
        Case "User defined text"
            For f = 1 To ID3.UDTextFrameNum
                lstFrames.AddItem ID3.UDTextFrameDescription(f) & ": " & ID3.UDTextFrame(f)
            Next f
        Case "Unique file identifier"
            For f = 1 To ID3.UFIDNum
                lstFrames.AddItem ID3.UFIDURL(f) & ": " & ID3.UFID(f)
            Next f
        Case "Link"
            For f = 1 To ID3.LinkFrameNum
                lstFrames.AddItem ID3.LinkFrameName(f) & ": " & ID3.LinkFrame(f)
            Next f
        Case "User defined link"
            For f = 1 To ID3.UDLinkFrameNum
                lstFrames.AddItem ID3.UDLinkFrameDescription(f) & ": " & ID3.UDLinkFrame(f)
            Next f
        Case "Involved people"
            For f = 1 To ID3.InvolvedPeopleNum
                lstFrames.AddItem ID3.InvolvedPeopleInvolvement(f) & ": " & ID3.InvolvedPeople(f)
            Next f
    End Select
End If
End Sub

Private Sub cmdNextPic_Click()
If lblPicNr < 20 Then lblPicNr = lblPicNr + 1
ID3.PictureID = lblPicNr
imgPicture.Picture = ID3.Picture
lblPicDesc = ID3.PictureDescription
End Sub

Private Sub cmdPrevPic_Click()
If lblPicNr > 0 Then lblPicNr = lblPicNr - 1
ID3.PictureID = lblPicNr
imgPicture.Picture = ID3.Picture
lblPicDesc = ID3.PictureDescription
End Sub

Private Sub cmdSave_Click()

If left(frmMain.SelectedFilename, 1) <> "C" Then
    Exit Sub
    Unload Me
End If

ID3.Filename = frmMain.SelectedFilename

ID3.Song = txtSong
ID3.Artist = txtArtist
ID3.Album = txtAlbum
ID3.Comment = txtComment
ID3.Genre = cboGenre
ID3.Year = txtYear
ID3.Track = txtTrack

frmMain.SongTitle = txtSong
frmMain.SongArtist = txtArtist

ID3.Save

Dim X
        
If Replace(frmMain.SongTitle, " ", "") = "" Then
    frmMain.SongTitle = frmMain.GetFileTitle(frmMain.SelectedFilename)
End If
        
If Replace(frmMain.SongArtist, " ", "") = "" Then
    frmMain.SongArtist = "Unknown"
End If
        
frmMain.RealScrollText = frmMain.SongArtist & " - " & frmMain.SongTitle
frmMain.ScrollText = frmMain.SongArtist & " - " & frmMain.SongTitle
        
For X = 1 To (frmMain.picFileContainer.ScaleWidth * 9.5) - Len(frmMain.ScrollText)
    frmMain.ScrollText = frmMain.ScrollText & " "
Next

frmMain.lstPlaylist.SelectedItem.ListSubItems(1).Text = frmMain.SongArtist
frmMain.lstPlaylist.SelectedItem.ListSubItems(2).Text = frmMain.SongTitle

Unload Me

End Sub

Private Sub ID3_Error(ErrNum As Integer)

    MsgBox ErrNum

End Sub

Private Sub Form_Load()

Dim Mode(0 To 3) As String, Emphasis(0 To 2) As String, tag(0 To 3) As String
Dim f As Integer

cboGenre.AddItem "[none]"
cboGenre.AddItem "A capella"
cboGenre.AddItem "Acid Jazz"
cboGenre.AddItem "Acid Punk"
cboGenre.AddItem "Acid"
cboGenre.AddItem "Acoustic"
cboGenre.AddItem "Alternative"
cboGenre.AddItem "AlternRock"
cboGenre.AddItem "Ambient"
cboGenre.AddItem "Avantgarde"
cboGenre.AddItem "Ballad"
cboGenre.AddItem "Bass"
cboGenre.AddItem "Bebob"
cboGenre.AddItem "Big Band"
cboGenre.AddItem "Bluegrass"
cboGenre.AddItem "Blues"
cboGenre.AddItem "Booty Bass"
cboGenre.AddItem "Cabaret"
cboGenre.AddItem "Celtic"
cboGenre.AddItem "Chamber Music"
cboGenre.AddItem "Chanson"
cboGenre.AddItem "Chorus"
cboGenre.AddItem "Christian Rap"
cboGenre.AddItem "Classic Rock"
cboGenre.AddItem "Classical"
cboGenre.AddItem "Club"
cboGenre.AddItem "Comedy"
cboGenre.AddItem "Country"
cboGenre.AddItem "Cult"
cboGenre.AddItem "Dance Hall"
cboGenre.AddItem "Dance"
cboGenre.AddItem "Darkwave"
cboGenre.AddItem "Death Metal"
cboGenre.AddItem "Disco"
cboGenre.AddItem "Dream"
cboGenre.AddItem "Drum Solo"
cboGenre.AddItem "Duet"
cboGenre.AddItem "Easy Listening"
cboGenre.AddItem "Electronic"
cboGenre.AddItem "Ethnic"
cboGenre.AddItem "Eurodance"
cboGenre.AddItem "Euro-House"
cboGenre.AddItem "Euro-Techno"
cboGenre.AddItem "Fast Fusion"
cboGenre.AddItem "Folk"
cboGenre.AddItem "Folklore"
cboGenre.AddItem "Folk-Rock"
cboGenre.AddItem "Freestyle"
cboGenre.AddItem "Funk"
cboGenre.AddItem "Fusion"
cboGenre.AddItem "Game"
cboGenre.AddItem "Gangsta"
cboGenre.AddItem "Gospel"
cboGenre.AddItem "Gothic Rock"
cboGenre.AddItem "Gothic"
cboGenre.AddItem "Grunge"
cboGenre.AddItem "Hard Rock"
cboGenre.AddItem "Hip-Hop"
cboGenre.AddItem "House"
cboGenre.AddItem "Humour"
cboGenre.AddItem "Industrial"
cboGenre.AddItem "Instrumental Pop"
cboGenre.AddItem "Instrumental Rock"
cboGenre.AddItem "Instrumental"
cboGenre.AddItem "Jazz"
cboGenre.AddItem "Jazz+Funk"
cboGenre.AddItem "Jungle"
cboGenre.AddItem "Latin"
cboGenre.AddItem "Lo-Fi"
cboGenre.AddItem "Meditative"
cboGenre.AddItem "Metal"
cboGenre.AddItem "Musical"
cboGenre.AddItem "National Folk"
cboGenre.AddItem "Native American"
cboGenre.AddItem "New Age"
cboGenre.AddItem "New Wave"
cboGenre.AddItem "Noise"
cboGenre.AddItem "Oldies"
cboGenre.AddItem "Opera"
cboGenre.AddItem "Other"
cboGenre.AddItem "Polka"
cboGenre.AddItem "Pop"
cboGenre.AddItem "Pop/Funk"
cboGenre.AddItem "Pop-Folk"
cboGenre.AddItem "Porn Groove"
cboGenre.AddItem "Power Ballad"
cboGenre.AddItem "Pranks"
cboGenre.AddItem "Primus"
cboGenre.AddItem "Progressive Rock"
cboGenre.AddItem "Psychadelic"
cboGenre.AddItem "Psychedelic Rock"
cboGenre.AddItem "Punk Rock"
cboGenre.AddItem "Punk"
cboGenre.AddItem "R&B"
cboGenre.AddItem "Rap"
cboGenre.AddItem "Rave"
cboGenre.AddItem "Reggae"
cboGenre.AddItem "Retro"
cboGenre.AddItem "Revival"
cboGenre.AddItem "Rhythmic Soul"
cboGenre.AddItem "Rock & Roll"
cboGenre.AddItem "Rock"
cboGenre.AddItem "Samba"
cboGenre.AddItem "Satire"
cboGenre.AddItem "Showtunes"
cboGenre.AddItem "Ska"
cboGenre.AddItem "Slow Jam"
cboGenre.AddItem "Slow Rock"
cboGenre.AddItem "Sonata"
cboGenre.AddItem "Soul"
cboGenre.AddItem "Sound Clip"
cboGenre.AddItem "Soundtrack"
cboGenre.AddItem "Southern Rock"
cboGenre.AddItem "Space"
cboGenre.AddItem "Speech"
cboGenre.AddItem "Swing"
cboGenre.AddItem "Symphonic Rock"
cboGenre.AddItem "Symphony"
cboGenre.AddItem "Tango"
cboGenre.AddItem "Techno"
cboGenre.AddItem "Techno-Industrial"
cboGenre.AddItem "Top 40"
cboGenre.AddItem "Trailer"
cboGenre.AddItem "Trance"
cboGenre.AddItem "Tribal"
cboGenre.AddItem "Trip-Hop"
cboGenre.AddItem "Vocal"
cboGenre.ListIndex = 0

cboFrames.AddItem "Text"
cboFrames.AddItem "User defined text"
cboFrames.AddItem "Unique file identifier"
cboFrames.AddItem "Link"
cboFrames.AddItem "User defined link"
cboFrames.AddItem "Involved people"
cboFrames.ListIndex = 0

Mode(0) = "Stereo"
Mode(1) = "Joint Stereo"
Mode(2) = "Dual Channel"
Mode(3) = "Mono"

Emphasis(0) = "None"
Emphasis(1) = "50/15 ms"
Emphasis(2) = "CCITT j.17"

tag(0) = "None"
tag(1) = "ID3v1"
tag(2) = "ID3v2"
tag(3) = "Both"

ID3.Filename = frmMain.SelectedFilename

ID3.Read

If Right(frmMain.SelectedFilename, 4) = ".mp3" Then

    lblInfo = "Layer: " & ID3.Layer & vbCrLf & _
            "MPEG version: " & ID3.MPEGVersion & vbCrLf & _
            "Bitrate: " & ID3.Bitrate & vbCrLf & _
            "Frequency: " & ID3.Frequency & vbCrLf & _
            "Private: " & ID3.PrivateBit & vbCrLf & _
            "Mode: " & Mode(ID3.Mode) & vbCrLf & _
            "Copyright: " & ID3.Copyright & vbCrLf & _
            "Original: " & ID3.Original & vbCrLf & _
            "Emphasis: " & Emphasis(ID3.Emphasis) & vbCrLf & _
            "Tags: " & tag(ID3.HaveTag) & vbCrLf & _
            "Length: " & frmMain.TimeToString(ID3.Length)

ElseIf Right(frmMain.SelectedFilename, 4) = ".avi" Then

    lblInfo = "NOT A TRUE MPEG FILE" & vbCrLf & _
            "OR THIS IS AN MPEG-4" & vbCrLf & _
            "FILE."
              
Else

    lblInfo = "Layer: " & ID3.Layer & vbCrLf & _
            "MPEG version: " & ID3.MPEGVersion & vbCrLf & _
            "Bitrate: " & ID3.Bitrate & vbCrLf & _
            "Frequency: " & ID3.Frequency & vbCrLf & _
            "Private: " & ID3.PrivateBit & vbCrLf & _
            "Mode: " & Mode(ID3.Mode) & vbCrLf & _
            "Copyright: " & ID3.Copyright & vbCrLf & _
            "Original: " & ID3.Original & vbCrLf & _
            "Emphasis: " & Emphasis(ID3.Emphasis) & vbCrLf & _
            "Tags: " & tag(ID3.HaveTag) & vbCrLf & _
            "Length: " & frmMain.TimeToString(frmMain.TotalTime)

End If

txtSong = ID3.Song
txtArtist = ID3.Artist
txtAlbum = ID3.Album
txtComment = ID3.Comment

If ID3.Genre = "" Then
    cboGenre.Text = "[none]"
Else
    cboGenre.Text = ID3.Genre
End If

txtYear = ID3.Year
txtTrack = ID3.Track


lblPicNr = 0
ID3.PictureID = 0
imgPicture = ID3.Picture
lblPicDesc = ID3.PictureDescription

cboFrames.ListIndex = 0

lstFrames.Clear
For f = 1 To ID3.TextFrameNum
    If ID3.TextFrameName(f) = "TCON" Then _
        lstFrames.AddItem ID3.TextFrameName(f) & " (decoded): " & ID3.DecodeTCON(ID3.TextFrame(f))
    lstFrames.AddItem ID3.TextFrameName(f) & ": " & ID3.TextFrame(f)
Next f

End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain.Enabled = True

End Sub

Private Sub txtTrack_KeyPress(KeyAscii As Integer)
If KeyAscii > 19 Then _
    If KeyAscii < 48 Or KeyAscii > 58 Then KeyAscii = 0
End Sub
