VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsID3v23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------
'clsID3v23 - reads ID3v2.3.x tags
'This Class was created by Paul Berlin 2002
'berlin_paul@hotmail.com
'This class uses the ID3v23x.DLL dll, Copyright (C) R. Glenn Scott
'----------------------------------------------------------

Option Explicit

Private Type tID3v2Data
  Track As String
  Title As String
  Artist As String
  Album As String
  Year As String
  Genre As String
  Comments As String
  Composer As String
  OrigArtist As String
  Copyright As String
  URL As String
  EncodedBy As String
End Type

Private ID3v2Info As tID3v2Data
Private strFilename As String
Private objTag As ID3v23x.clsID3v2

Public Property Get FileName() As String
  FileName = strFilename
End Property

Public Property Let FileName(ByVal sFilename As String)
  strFilename = sFilename
End Property

Public Function ReadTag() As Boolean
  On Error GoTo ErrH
  ReadTag = True
  
  Set objTag = New ID3v23x.clsID3v2
  
  objTag.ReadTag strFilename 'Reads tag
  If objTag.HasTag Then      'If there was a tag
    
    With objTag              'Fills ID3v2Info
      ID3v2Info.Album = Trim(.GetFrameValue(eAlbum))
      ID3v2Info.Artist = Trim(.GetFrameValue(eArtist))
      ID3v2Info.Comments = Trim(.GetFrameValue(eComment))
      ID3v2Info.Composer = Trim(.GetFrameValue(eComposer))
      ID3v2Info.Copyright = Trim(.GetFrameValue(eCopyright))
      ID3v2Info.EncodedBy = Trim(.GetFrameValue(eEncodedBy))
      ID3v2Info.Genre = Trim(.GetFrameValue(eGenre))
      ID3v2Info.OrigArtist = Trim(.GetFrameValue(eOrigArtist))
      ID3v2Info.Title = Trim(.GetFrameValue(eTitle))
      ID3v2Info.Track = Trim(.GetFrameValue(eTrack))
      ID3v2Info.URL = Trim(.GetFrameValue(eURL))
      ID3v2Info.Year = Trim(.GetFrameValue(eYear))
    End With
    
    Exit Function
  Else 'If there was no tag
    'Empty variables
    ID3v2Info.Album = ""
    ID3v2Info.Artist = ""
    ID3v2Info.Comments = ""
    ID3v2Info.Composer = ""
    ID3v2Info.Copyright = ""
    ID3v2Info.EncodedBy = ""
    ID3v2Info.Genre = ""
    ID3v2Info.OrigArtist = ""
    ID3v2Info.Title = ""
    ID3v2Info.Track = ""
    ID3v2Info.URL = ""
    ID3v2Info.Year = ""
    ReadTag = False
  End If
  
ErrH:
  ReadTag = False
End Function

Public Property Get tagTitle() As String
  tagTitle = ID3v2Info.Title
End Property

Public Property Let tagTitle(ByVal sTitle As String)
  ID3v2Info.Title = sTitle
End Property

Public Property Get tagArtist() As String
  tagArtist = ID3v2Info.Artist
End Property

Public Property Let tagArtist(ByVal sArtist As String)
  ID3v2Info.Artist = sArtist
End Property

Public Property Get tagAlbum() As String
  tagAlbum = ID3v2Info.Album
End Property

Public Property Let tagAlbum(ByVal sAlbum As String)
  ID3v2Info.Album = sAlbum
End Property

Public Property Get tagTrimmedGenre() As String
  'This returns the genre without the (#) in the beginning
  If Left(ID3v2Info.Genre, 1) = "(" Then
    tagTrimmedGenre = Right(ID3v2Info.Genre, Len(ID3v2Info.Genre) - InStr(1, ID3v2Info.Genre, ")"))
  Else
    tagTrimmedGenre = ID3v2Info.Genre
  End If
End Property

Public Property Get tagGenre() As String
  tagGenre = ID3v2Info.Genre
End Property

Public Property Let tagGenre(ByVal sGenre As String)
  ID3v2Info.Genre = sGenre
End Property

Public Property Get tagComments() As String
  tagComments = ID3v2Info.Comments
End Property

Public Property Let tagComments(ByVal sComments As String)
  ID3v2Info.Comments = sComments
End Property

Public Property Get tagTrack() As String
  tagTrack = ID3v2Info.Track
End Property

Public Property Let tagTrack(ByVal sTrack As String)
  ID3v2Info.Track = sTrack
End Property

Public Property Get tagYear() As String
  tagYear = ID3v2Info.Year
End Property

Public Property Let tagYear(ByVal sYear As String)
  ID3v2Info.Year = sYear
End Property

Public Function SaveTag() As Boolean
  On Error GoTo ErrH
  SaveTag = True
  
  Set objTag = New ID3v23x.clsID3v2
  
  With objTag           'sets new values from ID3v2Info
    .SetFrameValue eAlbum, ID3v2Info.Album
    .SetFrameValue eArtist, ID3v2Info.Artist
    .SetFrameValue eComment, ID3v2Info.Comments
    .SetFrameValue eComposer, ID3v2Info.Composer
    .SetFrameValue eCopyright, ID3v2Info.Copyright
    .SetFrameValue eEncodedBy, ID3v2Info.EncodedBy
    .SetFrameValue eGenre, ID3v2Info.Genre
    .SetFrameValue eOrigArtist, ID3v2Info.OrigArtist
    .SetFrameValue eTitle, ID3v2Info.Title
    .SetFrameValue eTrack, ID3v2Info.Track
    .SetFrameValue eURL, ID3v2Info.URL
    .SetFrameValue eYear, ID3v2Info.Year
    
    .WriteTag strFilename  'Writes
  End With
  
  
  Exit Function
ErrH:
  SaveTag = False
End Function

Public Function RemoveTag() As Boolean
  On Error GoTo ErrH
  RemoveTag = True
  
  Set objTag = New ID3v23x.clsID3v2
  objTag.RemoveTag strFilename

  Exit Function
ErrH:
  RemoveTag = False
End Function

Public Function HasTag() As Boolean
  On Error GoTo ErrH
  HasTag = True
  
  Set objTag = New ID3v23x.clsID3v2
  
  objTag.ReadTag strFilename 'Reads tag
  If Not objTag.HasTag Then   'If there wasn't a tag
    HasTag = False
  End If
    
  Exit Function
ErrH:
  HasTag = False
End Function

Public Property Get tagComposer() As String
  tagComposer = ID3v2Info.Composer
End Property

Public Property Let tagComposer(ByVal sComposer As String)
  ID3v2Info.Composer = sComposer
End Property

Public Property Get tagOrigArtist() As String
  tagOrigArtist = ID3v2Info.OrigArtist
End Property

Public Property Let tagOrigArtist(ByVal sOrigArtist As String)
  ID3v2Info.OrigArtist = sOrigArtist
End Property

Public Property Get tagCopyright() As String
  tagCopyright = ID3v2Info.Copyright
End Property

Public Property Let tagCopyright(ByVal sCopyright As String)
  ID3v2Info.Copyright = sCopyright
End Property

Public Property Get tagURL() As String
  tagURL = ID3v2Info.URL
End Property

Public Property Let tagURL(ByVal sURL As String)
  ID3v2Info.URL = sURL
End Property

Public Property Get tagEncodedBy() As String
  tagEncodedBy = ID3v2Info.EncodedBy
End Property

Public Property Let tagEncodedBy(ByVal sEncodedBy As String)
  ID3v2Info.EncodedBy = sEncodedBy
End Property

Private Sub Class_Terminate()
  Set objTag = Nothing
End Sub
