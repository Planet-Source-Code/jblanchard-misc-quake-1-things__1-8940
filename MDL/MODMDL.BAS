Attribute VB_Name = "Module1"
Option Explicit

Private Type vec3_t
    x   As Single
    y   As Single
    z   As Single
End Type
Private Type scalar_t
    scale   As Single
End Type

Private Type mdl_t
    id          As String * 4        ' 0x4F504449 = "IDPO" for IDPOLYGON
    Version     As Long         ' Version = 6
    scale       As vec3_t       ' Model scale factors.
    origin      As vec3_t       ' Model origin.
    radius      As scalar_t     ' Model bounding radius.
    offsets     As vec3_t       ' Eye position (useless?)
    numskins    As Long         ' the number of skin textures
    skinwidth   As Long         ' Width of skin texture
                                '           must be multiple of 8
    skinheight  As Long         ' Height of skin texture
                                '           must be multiple of 8
    numverts    As Long         ' Number of vertices
    numtris     As Long         ' Number of triangles surfaces
    numframes   As Long         ' Number of frames
    synctype    As Long         ' 0= synchron, 1= random
    flags       As Long         ' 0 (see Alias models)
    Size        As scalar_t     ' average size of triangles
End Type

Public ReadRecord As mdl_t

Public Sub Normalise(TriState As vec3_t, Start As Integer)
Dim xState As String, yState As String, zState As String
xState = Format$(TriState.x, "###.##")
yState = Format$(TriState.y, "###.##")
zState = Format$(TriState.z, "###.##")
If xState = "." Then xState = "0"
If yState = "." Then yState = "0"
If zState = "." Then zState = "0"
frmMain.txtOther(Start).Text = xState
frmMain.txtOther(Start + 1).Text = xState
frmMain.txtOther(Start + 2).Text = xState
End Sub


