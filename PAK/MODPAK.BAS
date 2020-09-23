Attribute VB_Name = "Module1"
Option Explicit

Private Type pakheader_t
    magic       As String * 4   ' Name of the new WAD format
    diroffset   As Long         ' Position of WAD directory from start of file
    dirsize     As Long         ' Number of entries * 0x40 (64 char)
End Type

Private Type pakentry_t
    filename    As String * 56  ' Name of the file, Unix style, with extension,
                                ' 50 chars, padded with '\0'.
    offset      As Long         ' Position of the entry in PACK file
    size   As Long              ' Size of the entry in PACK file
End Type

Public PakHeader As pakheader_t
Public PakEntry As pakentry_t
Public StoreEntry() As pakentry_t
