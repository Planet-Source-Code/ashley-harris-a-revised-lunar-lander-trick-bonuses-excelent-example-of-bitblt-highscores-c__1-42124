Attribute VB_Name = "globals"

'a list of pointers to the graphic objects, referenced by graphic name
'ie "rocket0002.msk.bmp" -> 346F5C6A45B5, and at that address in memory
'is the details of that picture
Public lib As New Dictionary

'provides access to the filesystem object, which allows a realiable
'way to manage files, ie managing the sprites and high score file
Public fso As New FileSystemObject

