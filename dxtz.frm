VERSION 5.00
Begin VB.Form frmDXTZ 
   BackColor       =   &H00000000&
   Caption         =   "DirectX Test Zone"
   ClientHeight    =   5595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   373
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   592
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   -30
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Use the right, left, up, and down arrow keys, plus page up/page down, to turn and move around the room."
      Top             =   5355
      Width           =   8895
   End
End
Attribute VB_Name = "frmDXTZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Pi = 3.1415927

'------------------------------------------------------------------
Dim X As Single 'VERY IMPORTANT VARIABLES!
Dim Y As Single 'Change these 3 variables to change the room size.
Dim Z As Single 'X=Width, Y=Height, Z=Depth
'------------------------------------------------------------------

'Requires the DirectX3 Type Library, available at Patrice Scribe's site -
'http://members.xoom.com/vba51/downloads/directx.zip

Dim D3DRM As IDirect3DRM
Dim Scene As IDirect3DRMFrame
Dim Camera As IDirect3DRMFrame
Dim CameraPos As D3DVECTOR
Dim CameraOrient As D3DVECTOR
Dim CameraDir As D3DVECTOR
Dim Clipper As IDirectDrawClipper
Dim Device As IDirect3DRMDevice
Dim Viewport As IDirect3DRMViewPort
Dim LightFrame As IDirect3DRMFrame
Dim Light As IDirect3DRMLight
Dim LightFrame2 As IDirect3DRMFrame
Dim Light2 As IDirect3DRMLight
Dim MeshBuilder As IDirect3DRMMeshBuilder

Private Sub Form_Load()
    X = 15 '-----------------------------------------
    Y = 3 'SET THE ROOM WIDTH HEIGHT AND DEPTH HERE.
    Z = 15 '-----------------------------------------
    ForeColor = vbGreen
    ' Create the Direct3D Retained Mode object
    Direct3DRMCreate D3DRM
    ' Create the scene frame
    D3DRM.CreateFrame Nothing, Scene
    ' Create the camera frame
    D3DRM.CreateFrame Scene, Camera
    ' Create the light frames
    ' INTERESTING NOTE - If you don't create extra light frames, you don't get wall shading!
    ' Check out the 'front' two corners - they have light.
    ' Check out the back two corners - their walls are shaded exactly the same.
    D3DRM.CreateFrame Scene, LightFrame
    D3DRM.CreateFrame Scene, LightFrame2
    ' Move the light frames to the two front  corners.
    LightFrame.SetPosition Scene, 1, Y / 2, Z - 1
    LightFrame2.SetPosition Scene, X - 1, Y / 2, Z - 1
    ' Points toward a corner (can't see spotlight if contained entirely on a face)
    LightFrame.SetOrientation Scene, 0.5, 0, 0.5, 0, 1, 0
    LightFrame2.SetOrientation Scene, 0.5, 0, 0.5, 0, 1, 0
    ' Create 2 spotlights
    D3DRM.CreateLightRGB D3DRMLIGHT_DIRECTIONAL, 0.5, 0.5, 0.5, Light
    D3DRM.CreateLightRGB D3DRMLIGHT_DIRECTIONAL, 0.2, 0.2, 0.2, Light2
    'Light.SetUmbra Pi / 4
    'Light.SetPenumbra Pi / 3
    Light.SetRange 15
    Light2.SetRange 15
    LightFrame.AddLight Light
    LightFrame2.AddLight Light2
    Scene.AddLight Light
    Scene.AddLight Light2
    ' Add ambient light to the scene
    Dim AmbientLight As IDirect3DRMLight
    D3DRM.CreateLightRGB D3DRMLIGHT_AMBIENT, 0.8, 0.8, 0.8, AmbientLight
    Scene.AddLight AmbientLight
    Set AmbientLight = Nothing
    ' Create a small world
    BuildWorld
    Camera.SetPosition Scene, X / 2, Y / 3, Z / 2
    'MeshBuilder.SetPerspective
End Sub

' Resize the viewport (fails if too large on my PC ?!)
Private Sub Form_Resize()
    Dim FormWidth As Long
    Dim FormHeight As Long
    Dim lblHeight As Long
    lblHeight = txtLabel.Height
    txtLabel.Move 0, (ScaleHeight - lblHeight), ScaleWidth, lblHeight
    FormWidth = ScaleWidth
    FormHeight = ScaleHeight - lblHeight 'Leave room for status at bottom.
    If Not (Viewport Is Nothing) Then Set Viewport = Nothing
    If Not (Clipper Is Nothing) Then Set Clipper = Nothing
    DirectDrawCreateClipper 0, Clipper, Nothing
    Clipper.SetHWnd 0, hWnd
    D3DRM.CreateDeviceFromClipper Clipper, ByVal 0&, FormWidth, FormHeight, Device
    D3DRM.CreateViewport Device, Camera, 0, 0, FormWidth, FormHeight, Viewport
    Device.SetDither 1
    Device.SetQuality D3DRMLIGHT_ON Or D3DRMFILL_SOLID Or D3DRMSHADE_GOURAUD
    UpdateScreen
End Sub

' Clear Direct3D Retained Mode objects
Private Sub Form_Unload(Cancel As Integer)
    Set Scene = Nothing
    Set Scene = Nothing
    Set Camera = Nothing
    Set Viewport = Nothing
    Set Device = Nothing
    Set D3DRM = Nothing
    Set Clipper = Nothing
End Sub

Sub BuildWorld()
    Dim FaceArray As IDirect3DRMFaceArray
    Dim Face As IDirect3DRMFace
    Dim Texture As IDirect3DRMTexture
    CreateBoxMesh X, Y, Z
    
    ' Perspective-corrected
    MeshBuilder.SetPerspective 1
    MeshBuilder.GetFaces FaceArray
       
    ' Textured front face
    FaceArray.GetElement 0, Face
    D3DRM.LoadTexture App.Path & "\stone.bmp", Texture
    Face.SetTexture Texture
    
    ' Textured back face
    FaceArray.GetElement 1, Face
    D3DRM.LoadTexture App.Path & "\stone.bmp", Texture
    Face.SetTexture Texture
    
    ' Textured right face
    FaceArray.GetElement 2, Face
    D3DRM.LoadTexture App.Path & "\stone.bmp", Texture
    Face.SetTexture Texture
    
    ' Textured left face
    FaceArray.GetElement 3, Face
    D3DRM.LoadTexture App.Path & "\stone.bmp", Texture
    Face.SetTexture Texture
    
    ' Texture top face
    FaceArray.GetElement 4, Face
    D3DRM.LoadTexture App.Path & "\floor.bmp", Texture
    Face.SetTexture Texture
    
    ' Texture bottom face
    FaceArray.GetElement 5, Face
    D3DRM.LoadTexture App.Path & "\floor2.bmp", Texture
    Face.SetTexture Texture
    
    Scene.AddVisual MeshBuilder
End Sub

' Camera movements
Private Static Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ShiftDown As Boolean
    Const Sin5 = 8.715574E-02!  ' Sin(5°)
    Const Cos5 = 0.9961947!     ' Cos(5°)
    
    On Error GoTo Form_KeyDown_Error
    
    ShiftDown = (Shift And vbShiftMask) > 0
    Select Case KeyCode
        ' Rotate left
        Case vbKeyLeft
            Camera.SetOrientation Camera, -Sin5, 0, Cos5, 0, 1, 0
        ' One step ahead (in the view direction)
        Case vbKeyUp
            If ShiftDown Then
                Camera.SetPosition Camera, 0, 0, 2 'go faster
            Else
                Camera.SetPosition Camera, 0, 0, 0.5
            End If
        ' Rotate right
        Case vbKeyRight
            Camera.SetOrientation Camera, Sin5, 0, Cos5, 0, 1, 0
        ' One step back
        Case vbKeyDown
            If ShiftDown Then
                Camera.SetPosition Camera, 0, 0, -2 'go faster
            Else
                Camera.SetPosition Camera, 0, 0, -0.5
            End If
        ' Look down (to fix)
        Case vbKeyPageDown
            Camera.SetOrientation Camera, 0, -Sin5, Cos5, 0, Cos5, 0
        ' Look up (to fix)
        Case vbKeyPageUp
            Camera.SetOrientation Camera, 0, Sin5, Cos5, 0, Cos5, 0
        ' Reset
        Case vbKeyEnd
            Camera.SetPosition Scene, X / 2, 1, Z / 2
        Case Else
            Exit Sub
    End Select
    UpdateScreen
    Exit Sub
' Track transient errors (where do they come from ?)
Form_KeyDown_Error:
    Debug.Print Now; " "; Err.Number; " "; Err.Description
    Resume
End Sub

Private Sub Form_Paint()
    UpdateScreen
End Sub

' Update screen display
Private Sub UpdateScreen()
    On Error GoTo UpdateScreen_Error
    Viewport.Clear
    Viewport.Render Scene
    Device.Update
    Exit Sub
UpdateScreen_Error:
    If Err.Number = 91 Then
        'viewport not set yet
    Else
        Debug.Print "UpdateScreen "; Now; " "; Err.Number; " "; Err.Description
    End If
End Sub

'-============================================================
' CreateBoxMesh
'-============================================================

Public Sub CreateBoxMesh(X As Single, Y As Single, Z As Single)

    Dim f As IDirect3DRMFace
    Dim tw As Single 'texture width
    Dim th As Single 'texture height
    tw = X 'taken from form_load - the width of the box
    th = Y 'taken from form_load - the height of the box
    
    D3DRM.CreateMeshBuilder MeshBuilder
    
    'Front Face
    D3DRM.CreateFace f
    f.AddVertex 0, Y, 0 'back showing
    f.AddVertex 0, 0, 0
    f.AddVertex X, 0, 0
    f.AddVertex X, Y, 0
    MeshBuilder.AddFace f
    
    'Back Face
    D3DRM.CreateFace f
    f.AddVertex X, Y, Z
    f.AddVertex X, 0, Z
    f.AddVertex 0, 0, Z
    f.AddVertex 0, Y, Z 'back showing
    MeshBuilder.AddFace f
    
    'Right face
    D3DRM.CreateFace f
    f.AddVertex X, Y, 0
    f.AddVertex X, 0, 0
    f.AddVertex X, 0, Z
    f.AddVertex X, Y, Z
    MeshBuilder.AddFace f

    'Left face
    D3DRM.CreateFace f
    f.AddVertex 0, 0, 0
    f.AddVertex 0, Y, 0
    f.AddVertex 0, Y, Z
    f.AddVertex 0, 0, Z
    MeshBuilder.AddFace f
    
    'Top face
    D3DRM.CreateFace f
    f.AddVertex X, Y, Z
    f.AddVertex 0, Y, Z
    f.AddVertex 0, Y, 0
    f.AddVertex X, Y, 0
    MeshBuilder.AddFace f
    
    'Bottom face
    D3DRM.CreateFace f
    f.AddVertex X, 0, 0
    f.AddVertex 0, 0, 0
    f.AddVertex 0, 0, Z
    f.AddVertex X, 0, Z
    MeshBuilder.AddFace f
    
    MeshBuilder.SetTextureCoordinates 3, 0, th 'Front face
    MeshBuilder.SetTextureCoordinates 2, 0, 0 'u controls flip
    MeshBuilder.SetTextureCoordinates 1, tw, 0 'v controls upsidedown
    MeshBuilder.SetTextureCoordinates 0, tw, th
    
    MeshBuilder.SetTextureCoordinates 7, 0, th 'Back face
    MeshBuilder.SetTextureCoordinates 6, 0, 0 'u controls flip
    MeshBuilder.SetTextureCoordinates 5, tw, 0 'v controls upsidedown
    MeshBuilder.SetTextureCoordinates 4, tw, th
    
    MeshBuilder.SetTextureCoordinates 11, 0, th 'Right face
    MeshBuilder.SetTextureCoordinates 10, 0, 0 'u controls flip
    MeshBuilder.SetTextureCoordinates 9, tw, 0  'v controls upsidedown
    MeshBuilder.SetTextureCoordinates 8, tw, th
    
    MeshBuilder.SetTextureCoordinates 15, tw, 0 'Left face
    MeshBuilder.SetTextureCoordinates 14, tw, th 'u controls flip
    MeshBuilder.SetTextureCoordinates 13, 0, th 'v controls upsidedown
    MeshBuilder.SetTextureCoordinates 12, 0, 0
    
    MeshBuilder.SetTextureCoordinates 19, tw, tw 'Top face
    MeshBuilder.SetTextureCoordinates 18, 0, tw 'u controls flip
    MeshBuilder.SetTextureCoordinates 17, 0, 0 'v controls upsidedown
    MeshBuilder.SetTextureCoordinates 16, tw, 0
                                
    MeshBuilder.SetTextureCoordinates 23, tw, tw 'Bottom face
    MeshBuilder.SetTextureCoordinates 22, 0, tw 'u controls flip
    MeshBuilder.SetTextureCoordinates 21, 0, 0 'v controls upsidedown
    MeshBuilder.SetTextureCoordinates 20, tw, 0
                                
    MeshBuilder.SetName "Box"
    
End Sub

Private Sub txtLabel_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub
