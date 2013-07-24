Attribute VB_Name = "modRendering"

Option Explicit

' DirectX8 Object
Private Directx8 As Directx8 ' The master DirectX object.
Private Direct3D As Direct3D8 ' Controls all things 3D.
Public Direct3D_Device As Direct3DDevice8 ' Represents the hardware rendering.
Private Direct3DX As D3DX8

' The 2D (Transformed and Lit) vertex format.
Private Const FVF_TLVERTEX As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE

' The 2D (Transformed and Lit) vertex format type.
Public Type TLVERTEX
    x As Single
    y As Single
    Z As Single
    RHW As Single
    Color As Long
    TU As Single
    TV As Single
End Type

Private Vertex_List(3) As TLVERTEX ' 4 vertices will make a square.

' Some color depth constants to help make the DX constants more readable.
Private Const COLOR_DEPTH_16_BIT As Long = D3DFMT_R5G6B5
Private Const COLOR_DEPTH_24_BIT As Long = D3DFMT_A8R8G8B8
Private Const COLOR_DEPTH_32_BIT As Long = D3DFMT_X8R8G8B8

Public RenderingMode As Long

Private Direct3D_Window As D3DPRESENT_PARAMETERS ' Backbuffer and viewport description.
Private Display_Mode As D3DDISPLAYMODE

' Graphic Textures
Public Tex_Item() As DX8TextureRec ' Arrays
Public Tex_Character() As DX8TextureRec
Public Tex_Paperdoll() As DX8TextureRec
Public Tex_Tileset() As DX8TextureRec
Public Tex_Resource() As DX8TextureRec
Public Tex_Animation() As DX8TextureRec
Public Tex_SpellIcon() As DX8TextureRec
Public Tex_Face() As DX8TextureRec
Public Tex_Fog() As DX8TextureRec
Public Tex_Panorama() As DX8TextureRec
Public Tex_Emoticon() As DX8TextureRec
Public Tex_Blood As DX8TextureRec ' Singles
Public Tex_Misc As DX8TextureRec
Public Tex_Direction As DX8TextureRec
Public Tex_Target As DX8TextureRec
Public Tex_Bars As DX8TextureRec
Public Tex_Selection As DX8TextureRec
Public Tex_White As DX8TextureRec
Public Tex_Weather As DX8TextureRec
Public Tex_ChatBubble As DX8TextureRec
Public Tex_Fade As DX8TextureRec
Public Tex_Equip As DX8TextureRec
Public Tex_Base As DX8TextureRec

' Character Editor Sprite
Public Tex_CharSprite As DX8TextureRec
Public LastCharSpriteTimer As Long
Private CharSpritePos As Byte

' Number of graphic files
Public NumTileSets As Long
Public NumCharacters As Long
Public NumPaperdolls As Long
Public NumItems As Long
Public NumResources As Long
Public NumAnimations As Long
Public NumSpellIcons As Long
Public NumFaces As Long
Public NumFogs As Long
Public NumPanoramas As Long
Public NumEmoticons As Long

Public Type DX8TextureRec
    Texture As Long
    Width As Long
    Height As Long
    filepath As String
    TexWidth As Long
    TexHeight As Long
    ImageData() As Byte
    HasData As Boolean
End Type

Public Type GlobalTextureRec
    Texture As Direct3DTexture8
    TexWidth As Long
    TexHeight As Long
End Type

Public Type RECT
    Top As Long
    Left As Long
    Bottom As Long
    Right As Long
End Type

Public gTexture() As GlobalTextureRec
Public NumTextures As Long

' ********************
' ** Initialization **
' ********************
Public Function InitDX8() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Directx8 = New Directx8 ' Creates the DirectX object.
    Set Direct3D = Directx8.Direct3DCreate() ' Creates the Direct3D object using the DirectX object.
    Set Direct3DX = New D3DX8
    
    Direct3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display_Mode ' Use the current display mode that you are on (resolution).
    Direct3D_Window.Windowed = True ' The app will be in windowed mode.
    
    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_DISCARD ' Refresh when the monitor does.
    Direct3D_Window.BackBufferFormat = Display_Mode.Format ' Sets the format that was retrieved into the backbuffer.
    
    ' Creates the rendering device with some useful info, along with the info
    ' DispMode.Format = D3DFMT_X8R8G8B8
    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_COPY
    Direct3D_Window.BackBufferCount = 1 '1 backbuffer only
    Direct3D_Window.BackBufferWidth = frmMain.picScreen.ScaleWidth 'Match the backbuffer width with the display width
    Direct3D_Window.BackBufferHeight = frmMain.picScreen.ScaleHeight 'Match the backbuffer height with the display height
    Direct3D_Window.hDeviceWindow = frmMain.picScreen.hwnd 'Use frmMain as the device window.
    
    ' We've already setup for Direct3D_Window.
    If TryCreateDirectX8Device = False Then
        MsgBox "Unable to initialize DirectX8. You may be missing dx8vb.dll or have incompatible hardware to use DirectX8."
        DestroyGame
    End If

    With Direct3D_Device
        .SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
    
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False
        
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
    End With
    
    ' Initialize the surfaces
    LoadTextures
    
    ' We're done
    InitDX8 = True
    Exit Function
    
' Error handler
errorhandler:
    HandleError "InitDX8", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Function Ceiling(dblValIn As Double, dblCeilIn As Double) As Double
    ' Round it
    Ceiling = Round(dblValIn / dblCeilIn, 0) * dblCeilIn
    
    ' If it rounded down, force it up
    If Ceiling < dblValIn Then Ceiling = Ceiling + dblCeilIn
End Function

Public Sub DestroyDX8()
    UnloadTextures
    
    Set Direct3DX = Nothing
    Set Direct3D_Device = Nothing
    Set Direct3D = Nothing
    Set Directx8 = Nothing
End Sub

Public Sub DrawGDI()
    ' Cycle through in-game stuff before cycling through editors
    If frmMenu.Visible Then
        If frmMenu.picCharacter.Visible Then NewCharacterDrawSprite
    End If
    
    If frmMain.Visible Then
        If frmMain.picTempInv.Visible Then DrawDraggedItem frmMain.picTempInv.Left, frmMain.picTempInv.Top
        If frmMain.picTempSpell.Visible Then DrawDraggedSpell frmMain.picTempSpell.Left, frmMain.picTempSpell.Top
        If frmMain.picSpellDesc.Visible Then DrawSpellDesc LastSpellDesc
        If frmMain.picItemDesc.Visible Then DrawItemDesc LastItemDesc
        If frmMain.picHotbar.Visible Then DrawHotbar
        If frmMain.picInventory.Visible Then DrawInventory
        If frmMain.picCharacter.Visible Then DrawPlayerCharFace
        If frmMain.picEquipment.Visible Then DrawEquipment
        If frmMain.picChatFace.Visible Then DrawEventChatFace
        If frmMain.picSpells.Visible Then DrawPlayerSpells
        If frmMain.picShop.Visible Then DrawShop
        If frmMain.picTempBank.Visible Then DrawBankItem frmMain.picTempBank.Left, frmMain.picTempBank.Top
        If frmMain.picBank.Visible Then DrawBank
        If frmMain.picTrade.Visible Then DrawTrade
    End If
    
    If frmEditor_Animation.Visible Then
        EditorAnim_DrawAnim
    End If
    
    If frmEditor_Item.Visible Then
        EditorItem_DrawItem
        EditorItem_DrawPaperdoll
    End If
    
    If frmEditor_Map.Visible Then
        EditorMap_DrawTileset
        If frmEditor_Map.fraMapItem.Visible Then EditorMap_DrawMapItem
    End If
    
    ' Renders random tiles in map editor
    If frmEditor_Map.chkRandom.Value = 1 Then
        Call EditorMap_DrawRandom
    End If
    
    ' Character editor
    If frmCharEditor.Visible And Tex_CharSprite.Texture > 0 And requestedPlayer.Sprite > 0 Then
        If LastCharSpriteTimer + 300 < timeGetTime Then
            LastCharSpriteTimer = timeGetTime
            Call EditorChar_AnimSprite
        End If
    End If

    If frmEditor_MapProperties.Visible Then
        EditorMapProperties_DrawPanorama
    End If
    
    If frmEditor_NPC.Visible Then
        EditorNPC_DrawSprite
    End If
    
    If frmEditor_Resource.Visible Then
        EditorResource_DrawSprite
    End If
    
    If frmEditor_Spell.Visible Then
        EditorSpell_DrawIcon
    End If
    
    If frmEditor_Events.Visible Then
        EditorEvent_DrawFace
        EditorEvent_DrawFace2
        EditorEvent_DrawGraphic
    End If
    
    If frmEditor_Emoticon.Visible Then
        EditorEmoticon_DrawIcon
    End If
    
    If frmEditor_Class.Visible Then
        With frmEditor_Class
            If .scrlMSprite.Visible Then
                Call EditorClass_DrawSprite(0)
            Else
                Call EditorClass_DrawSprite(1)
            End If
            
            If .scrlMFace.Visible Then
                Call EditorClass_DrawFace(0)
            Else
                Call EditorClass_DrawFace(1)
            End If
        End With
    End If
End Sub

Function TryCreateDirectX8Device() As Boolean
    Dim i As Long

    On Error GoTo nexti
    
    For i = 1 To 4
        Select Case i
            Case 1
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.picScreen.hwnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            Case 2
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.picScreen.hwnd, D3DCREATE_MIXED_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            Case 3
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.picScreen.hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            Case 4
                TryCreateDirectX8Device = False
                Exit Function
        End Select
nexti:
    Next

End Function

Function GetNearestPOT(Value As Long) As Long
    Dim i As Long

    Do While 2 ^ i < Value
        i = i + 1
    Loop
    
    GetNearestPOT = 2 ^ i
End Function

Public Sub LoadTexture(ByRef TextureRec As DX8TextureRec)
    Dim SourceBitmap As cGDIpImage, ConvertedBitmap As cGDIpImage, GDIGraphics As cGDIpRenderer, GDIToken As cGDIpToken, i As Long
    Dim newWidth As Long, newHeight As Long, ImageData() As Byte, fn As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If TextureRec.HasData = False Then
        Set GDIToken = New cGDIpToken
        
        ' Make sure it loaded correctly
        If GDIToken.Token = 0& Then MsgBox "GDI+ failed to load, exiting game!": DestroyGame
        
        Set SourceBitmap = New cGDIpImage
        Call SourceBitmap.LoadPicture_FileName(TextureRec.filepath, GDIToken)
        
        TextureRec.Width = SourceBitmap.Width
        TextureRec.Height = SourceBitmap.Height
        
        newWidth = GetNearestPOT(TextureRec.Width)
        newHeight = GetNearestPOT(TextureRec.Height)
        
        If newWidth <> SourceBitmap.Width Or newHeight <> SourceBitmap.Height Then
            Set ConvertedBitmap = New cGDIpImage
            Set GDIGraphics = New cGDIpRenderer
            i = GDIGraphics.CreateGraphicsFromImageClass(SourceBitmap)
            Call ConvertedBitmap.LoadPicture_FromNothing(newHeight, newWidth, i, GDIToken) 'I HAVE NO IDEA why this is backwards but it works.
            Call GDIGraphics.DestroyHGraphics(i)
            i = GDIGraphics.CreateGraphicsFromImageClass(ConvertedBitmap)
            Call GDIGraphics.AttachTokenClass(GDIToken)
            Call GDIGraphics.RenderImageClassToHGraphics(SourceBitmap, i)
            Call ConvertedBitmap.SaveAsPNG(ImageData)
            GDIGraphics.DestroyHGraphics (i)
            TextureRec.ImageData = ImageData
            Set ConvertedBitmap = Nothing
            Set GDIGraphics = Nothing
            Set SourceBitmap = Nothing
        Else
            Call SourceBitmap.SaveAsPNG(ImageData)
            TextureRec.ImageData = ImageData
            Set SourceBitmap = Nothing
        End If
    Else
        ImageData = TextureRec.ImageData
    End If
    
    Set gTexture(TextureRec.Texture).Texture = Direct3DX.CreateTextureFromFileInMemoryEx(Direct3D_Device, _
        ImageData(0), _
        UBound(ImageData) + 1, _
        newWidth, _
        newHeight, _
        D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, ByVal (0), ByVal 0, ByVal 0)
    
    gTexture(TextureRec.Texture).TexWidth = newWidth
    gTexture(TextureRec.Texture).TexHeight = newHeight
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "LoadTexture", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub LoadTextures()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call CheckTilesets
    Call CheckCharacters
    Call CheckPaperdolls
    Call CheckAnimations
    Call CheckItems
    Call CheckResources
    Call CheckSpellIcons
    Call CheckFaces
    Call CheckFogs
    Call CheckPanoramas
    Call CheckEmoticons
    
    NumTextures = NumTextures + 13
    
    ReDim Preserve gTexture(NumTextures)
    Tex_Base.filepath = App.Path & "\data files\graphics\gui\main\base.png"
    Tex_Base.Texture = NumTextures - 11
    LoadTexture Tex_Base
    Tex_Equip.filepath = App.Path & "\data files\graphics\gui\main\equip.png"
    Tex_Equip.Texture = NumTextures - 10
    LoadTexture Tex_Equip
    Tex_Fade.filepath = App.Path & "\data files\graphics\misc\fader.png"
    Tex_Fade.Texture = NumTextures - 9
    LoadTexture Tex_Fade
    Tex_ChatBubble.filepath = App.Path & "\data files\graphics\misc\chatbubble.png"
    Tex_ChatBubble.Texture = NumTextures - 8
    LoadTexture Tex_ChatBubble
    Tex_Weather.filepath = App.Path & "\data files\graphics\misc\weather.png"
    Tex_Weather.Texture = NumTextures - 7
    LoadTexture Tex_Weather
    Tex_White.filepath = App.Path & "\data files\graphics\misc\white.png"
    Tex_White.Texture = NumTextures - 6
    LoadTexture Tex_White
    Tex_Direction.filepath = App.Path & "\data files\graphics\misc\direction.png"
    Tex_Direction.Texture = NumTextures - 5
    LoadTexture Tex_Direction
    Tex_Target.filepath = App.Path & "\data files\graphics\misc\target.png"
    Tex_Target.Texture = NumTextures - 4
    LoadTexture Tex_Target
    Tex_Misc.filepath = App.Path & "\data files\graphics\misc\misc.png"
    Tex_Misc.Texture = NumTextures - 3
    LoadTexture Tex_Misc
    Tex_Blood.filepath = App.Path & "\data files\graphics\misc\blood.png"
    Tex_Blood.Texture = NumTextures - 2
    LoadTexture Tex_Blood
    Tex_Bars.filepath = App.Path & "\data files\graphics\misc\bars.png"
    Tex_Bars.Texture = NumTextures - 1
    LoadTexture Tex_Bars
    Tex_Selection.filepath = App.Path & "\data files\graphics\misc\select.png"
    Tex_Selection.Texture = NumTextures
    LoadTexture Tex_Selection
    
    EngineInitFontTextures
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "LoadTextures", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub UnloadTextures()
Dim i As Long
    
    ' If debug mode, handle error then exit out
    On Error Resume Next
    
    For i = 1 To NumTextures
        Set gTexture(i).Texture = Nothing
        ZeroMemory ByVal VarPtr(gTexture(i)), LenB(gTexture(i))
    Next
    
    ReDim gTexture(1)

    For i = 1 To NumTileSets
        Tex_Tileset(i).Texture = 0
    Next

    For i = 1 To NumItems
        Tex_Item(i).Texture = 0
    Next

    For i = 1 To NumCharacters
        Tex_Character(i).Texture = 0
    Next
    
    For i = 1 To NumPaperdolls
        Tex_Paperdoll(i).Texture = 0
    Next
    
    For i = 1 To NumResources
        Tex_Resource(i).Texture = 0
    Next
    
    For i = 1 To NumAnimations
        Tex_Animation(i).Texture = 0
    Next
    
    For i = 1 To NumSpellIcons
        Tex_SpellIcon(i).Texture = 0
    Next
    
    For i = 1 To NumFaces
        Tex_Face(i).Texture = 0
    Next
    
    For i = 1 To NumPanoramas
        Tex_Panorama(i).Texture = 0
    Next
    
    For i = 1 To NumEmoticons
        Tex_Emoticon(i).Texture = 0
    Next

    Tex_Equip.Texture = 0
    Tex_Fade.Texture = 0
    Tex_ChatBubble.Texture = 0
    Tex_Weather.Texture = 0
    Tex_White.Texture = 0
    Tex_Bars.Texture = 0
    Tex_Misc.Texture = 0
    Tex_Blood.Texture = 0
    Tex_Direction.Texture = 0
    Tex_Target.Texture = 0
    Tex_Selection.Texture = 0
    
    UnloadFontTextures
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "UnloadTextures", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' **************
' ** Drawing **
' **************
Public Sub RenderTexture(ByRef TextureRec As DX8TextureRec, ByVal dX As Single, ByVal dY As Single, ByVal sX As Single, ByVal sY As Single, ByVal dWidth As Single, ByVal dHeight As Single, ByVal sWidth As Single, ByVal sHeight As Single, Optional Color As Long = -1)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Dim textureNum As Long
    Dim textureWidth As Long, textureHeight As Long, sourceX As Single, sourceY As Single, sourceWidth As Single, sourceHeight As Single
    textureNum = TextureRec.Texture
    
    textureWidth = gTexture(textureNum).TexWidth
    textureHeight = gTexture(textureNum).TexHeight
    
    If sY + sHeight > textureHeight Then Exit Sub
    If sX + sWidth > textureWidth Then Exit Sub
    If sX < 0 Then Exit Sub
    If sY < 0 Then Exit Sub

    sX = sX - 0.5
    sY = sY - 0.5
    dY = dY - 0.5
    dX = dX - 0.5
    sWidth = sWidth
    sHeight = sHeight
    dWidth = dWidth
    dHeight = dHeight
    sourceX = (sX / textureWidth)
    sourceY = (sY / textureHeight)
    sourceWidth = ((sX + sWidth) / textureWidth)
    sourceHeight = ((sY + sHeight) / textureHeight)
    
    Vertex_List(0) = Create_TLVertex(dX, dY, 0, 1, Color, 0, sourceX + 0.000003, sourceY + 0.000003)
    Vertex_List(1) = Create_TLVertex(dX + dWidth, dY, 0, 1, Color, 0, sourceWidth + 0.000003, sourceY + 0.000003)
    Vertex_List(2) = Create_TLVertex(dX, dY + dHeight, 0, 1, Color, 0, sourceX + 0.000003, sourceHeight + 0.000003)
    Vertex_List(3) = Create_TLVertex(dX + dWidth, dY + dHeight, 0, 1, Color, 0, sourceWidth + 0.000003, sourceHeight + 0.000003)
    
    Direct3D_Device.SetTexture 0, gTexture(textureNum).Texture
    Direct3D_Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex_List(0), Len(Vertex_List(0))
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "RenderTexture", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub RenderTextureByRects(TextureRec As DX8TextureRec, sRect As RECT, dRect As RECT)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    RenderTexture TextureRec, dRect.Left, dRect.Top, sRect.Left, sRect.Top, dRect.Right - dRect.Left, dRect.Bottom - dRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "RenderTextureByRects", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' This function will make it much easier to setup the vertices with the info it needs.
Private Function Create_TLVertex(x As Single, y As Single, Z As Single, RHW As Single, Color As Long, Specular As Long, TU As Single, TV As Single) As TLVERTEX
    Create_TLVertex.x = x
    Create_TLVertex.y = y
    Create_TLVertex.Z = Z
    Create_TLVertex.RHW = RHW
    Create_TLVertex.Color = Color
    Create_TLVertex.TU = TU
    Create_TLVertex.TV = TV
End Function

Public Sub DrawGrid(ByVal x As Long, ByVal y As Long)
    Dim Top As Long, Left As Long
    
    ' Render grid
    Top = 24
    Left = 0

    RenderTexture Tex_Direction, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), Left, Top, 32, 32, 32, 32
End Sub

' Directional blocking
Public Sub DrawDirection(ByVal x As Long, ByVal y As Long)
    Dim i As Long, Top As Long, Left As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Render dir blobs
    For i = 1 To 4
        Left = (i - 1) * 8
        
        ' Find out whether render blocked or not
        If Not IsDirBlocked(Map.Tile(x, y).DirBlock, CByte(i)) Then
            Top = 8
        Else
            Top = 16
        End If
       
        RenderTexture Tex_Direction, ConvertMapX(x * PIC_X) + DirArrowX(i), ConvertMapY(y * PIC_Y) + DirArrowY(i), Left, Top, 8, 8, 8, 8
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawDirection", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawTarget(ByVal x As Long, ByVal y As Long)
    Dim sRect As RECT
    Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Tex_Target.Texture = 0 Then Exit Sub
    
    Width = Tex_Target.Width / 2
    Height = Tex_Target.Height

    With sRect
        .Top = 0
        .Bottom = Height
        .Left = 0
        .Right = Width
    End With
    
    x = x - ((Width - 32) / 2)
    y = y - (Height / 2)
    
    x = ConvertMapX(x)
    y = ConvertMapY(y)
    
    ' Clipping
    If y < 0 Then
        With sRect
            .Top = .Top - y
        End With
        y = 0
    End If

    If x < 0 Then
        With sRect
            .Left = .Left - x
        End With
        x = 0
    End If
    
    RenderTexture Tex_Target, x, y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawTarget", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawHover(ByVal tType As Long, ByVal Target As Long, ByVal x As Long, ByVal y As Long)
    Dim sRect As RECT
    Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Tex_Target.Texture = 0 Then Exit Sub
    
    Width = Tex_Target.Width / 2
    Height = Tex_Target.Height

    With sRect
        .Top = 0
        .Bottom = Height
        .Left = Width
        .Right = .Left + Width
    End With
    
    x = x - ((Width - 32) / 2)
    y = y - (Height / 2)

    x = ConvertMapX(x)
    y = ConvertMapY(y)
    
    ' Clipping
    If y < 0 Then
        With sRect
            .Top = .Top - y
        End With
        y = 0
    End If

    If x < 0 Then
        With sRect
            .Left = .Left - x
        End With
        x = 0
    End If
    
    RenderTexture Tex_Target, x, y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawHover", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawMapLowerTiles(ByVal x As Long, ByVal y As Long)
    Dim rec As RECT
    Dim i As Long, Alpha As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Map.Tile(x, y)
        For i = MapLayer.Ground To MapLayer.Cover
            If i < CurrentLayer And frmEditor_Map.ChkDimLayers = 1 And InMapEditor Then
                Alpha = 255 - ((CurrentLayer - i) * 48)
            Else
                Alpha = 255
            End If
            
            If Autotile(x, y).Layer(i).RenderState = RENDER_STATE_NORMAL Then
                ' Draw normally
                RenderTexture Tex_Tileset(.Layer(i).Tileset), ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), .Layer(i).x * 32, .Layer(i).y * 32, 32, 32, 32, 32, D3DColorARGB(Alpha, 255, 255, 255)
            ElseIf Autotile(x, y).Layer(i).RenderState = RENDER_STATE_AUTOTILE And Options.Autotile = 1 Then
                ' Draw autotiles
                DrawAutoTile i, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), 1, x, y, Alpha
                DrawAutoTile i, ConvertMapX((x * PIC_X) + 16), ConvertMapY(y * PIC_Y), 2, x, y, Alpha
                DrawAutoTile i, ConvertMapX(x * PIC_X), ConvertMapY((y * PIC_Y) + 16), 3, x, y, Alpha
                DrawAutoTile i, ConvertMapX((x * PIC_X) + 16), ConvertMapY((y * PIC_Y) + 16), 4, x, y, Alpha
            End If
        Next
    End With
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawMapLowerTiles", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawMapUpperTiles(ByVal x As Long, ByVal y As Long)
    Dim rec As RECT
    Dim i As Long, Alpha As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Map.Tile(x, y)
        For i = MapLayer.Fringe To MapLayer.Roof
            If i < CurrentLayer And frmEditor_Map.ChkDimLayers = 1 And InMapEditor Then
                Alpha = 255 - ((CurrentLayer - i) * 48)
            Else
                Alpha = 255
            End If
            
            If Autotile(x, y).Layer(i).RenderState = RENDER_STATE_NORMAL Then
                ' Draw normally
                RenderTexture Tex_Tileset(.Layer(i).Tileset), ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), .Layer(i).x * 32, .Layer(i).y * 32, 32, 32, 32, 32, D3DColorARGB(Alpha, 255, 255, 255)
            ElseIf Autotile(x, y).Layer(i).RenderState = RENDER_STATE_AUTOTILE And Options.Autotile = 1 Then
                ' Draw autotiles
                DrawAutoTile i, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), 1, x, y, Alpha
                DrawAutoTile i, ConvertMapX((x * PIC_X) + 16), ConvertMapY(y * PIC_Y), 2, x, y, Alpha
                DrawAutoTile i, ConvertMapX(x * PIC_X), ConvertMapY((y * PIC_Y) + 16), 3, x, y, Alpha
                DrawAutoTile i, ConvertMapX((x * PIC_X) + 16), ConvertMapY((y * PIC_Y) + 16), 4, x, y, Alpha
            End If
        Next

        ' Tile preview
        If InMapEditor Then
            If frmEditor_Map.chkTilePreview.Value And frmEditor_Map.chkRandom = 0 And frmEditor_Map.scrlAutotile.Value = 0 And frmEditor_Map.OptLayers.Value Then
                Call EditorMap_DrawTilePreview
            End If
        End If
    End With
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawMapUpperTiles", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawBlood(ByVal Index As Long)
    Dim rec As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Load blood
    BloodCount = Tex_Blood.Width / 32
    
    With Blood(Index)
        If .Alpha <= 0 Then
            Call ClearBlood(Index)
            Exit Sub
        End If
        
        ' Check if we should be seeing it
        If .Timer + 20000 < timeGetTime Then
            .Alpha = .Alpha - 1
        End If
        
        rec.Top = 0
        rec.Bottom = PIC_Y
        rec.Left = (.Sprite - 1) * PIC_X
        rec.Right = rec.Left + PIC_X
        RenderTexture Tex_Blood, ConvertMapX(.x * PIC_X), ConvertMapY(.y * PIC_Y), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorARGB(Blood(Index).Alpha, 255, 255, 255)
    End With
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawBlood", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawAnimation(ByVal Index As Long, ByVal Layer As Long)
    Dim Sprite As Long
    Dim sRect As RECT
    Dim dRect As RECT
    Dim i As Long
    Dim Width As Long, Height As Long
    Dim LoopTime As Long
    Dim FrameCount As Long
    Dim x As Long, y As Long
    Dim lockIndex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If AnimInstance(Index).Animation = 0 Then
        ClearAnimInstance Index
        Exit Sub
    End If
    
    Sprite = Animation(AnimInstance(Index).Animation).Sprite(Layer)
    
    ' Make sure the sprite exists
    If Sprite < 1 Or Sprite > NumAnimations Then Exit Sub
    
    FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
    
    ' Total width divided by frame count
    Width = Tex_Animation(Sprite).Width / FrameCount
    Height = Tex_Animation(Sprite).Height
    
    sRect.Top = 0
    sRect.Bottom = Height
    sRect.Left = (AnimInstance(Index).frameIndex(Layer) - 1) * Width
    sRect.Right = sRect.Left + Width
    
    ' Change x or y if locked
    If AnimInstance(Index).LockType > TARGET_TYPE_NONE Then ' if <> none
        ' Is a player
        If AnimInstance(Index).LockType = TARGET_TYPE_PLAYER Then
            ' Quick save the index
            lockIndex = AnimInstance(Index).lockIndex
            
            ' Check if is ingame
            If IsPlaying(lockIndex) Then
                ' Check if on same map
                If GetPlayerMap(lockIndex) = GetPlayerMap(MyIndex) Then
                    ' Is on map, is playing, set x & y
                    x = (GetPlayerX(lockIndex) * PIC_X) + 16 - (Width / 2) + TempPlayer(lockIndex).xOffset
                    y = (GetPlayerY(lockIndex) * PIC_Y) + 16 - (Height / 2) + TempPlayer(lockIndex).yOffset
                End If
            End If
        ElseIf AnimInstance(Index).LockType = TARGET_TYPE_NPC Then
            ' Quick save the index
            lockIndex = AnimInstance(Index).lockIndex
            
            ' Check if NPC exists
            If MapNPC(lockIndex).Num > 0 Then
                ' Check if alive
                If MapNPC(lockIndex).Vital(Vitals.HP) > 0 Then
                    ' Exists, is alive, set x & y
                    x = (MapNPC(lockIndex).x * PIC_X) + 16 - (Width / 2) + MapNPC(lockIndex).xOffset
                    y = (MapNPC(lockIndex).y * PIC_Y) + 16 - (Height / 2) + MapNPC(lockIndex).yOffset
                Else
                    ' Npc not alive anymore, kill the animation
                    ClearAnimInstance Index
                    Exit Sub
                End If
            Else
                ' Npc not alive anymore, kill the animation
                ClearAnimInstance Index
                Exit Sub
            End If
        End If
    Else
        ' No lock, default x + y
        x = (AnimInstance(Index).x * 32) + 16 - (Width / 2)
        y = (AnimInstance(Index).y * 32) + 16 - (Height / 2)
    End If
    
    x = ConvertMapX(x)
    y = ConvertMapY(y)

    ' Clip to screen
    If y < 0 Then
        With sRect
            .Top = .Top - y
        End With

        y = 0
    End If

    If x < 0 Then
        With sRect
            .Left = .Left - x
        End With

        x = 0
    End If
    
    RenderTexture Tex_Animation(Sprite), x, y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
    Exit Sub
        
' Error handler
errorhandler:
    HandleError "DrawAnimation", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawMapItem(ByVal ItemNum As Long)
    Dim PicNum As Integer, x As Long, i As Long
    Dim rec As RECT
    Dim MaxFrames As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    x = 0
    
    ' If it's not ours then don't render
    If x = 0 Then
        If Not Trim$(MapItem(ItemNum).PlayerName) = GetPlayerName(MyIndex) Then
            If Not Trim$(MapItem(ItemNum).PlayerName) = vbNullString Then Exit Sub
        End If
    End If

    PicNum = Item(MapItem(ItemNum).Num).Pic

    If PicNum < 1 Or PicNum > NumItems Then Exit Sub
    
    If Tex_Item(PicNum).Width > 64 Then ' Has more than 1 frame
        With rec
            .Top = 0
            .Bottom = 32
            .Left = (MapItem(ItemNum).Frame * 32)
            .Right = .Left + 32
        End With
    Else
        With rec
            .Top = 0
            .Bottom = PIC_Y
            .Left = 0
            .Right = PIC_X
        End With
    End If
    
    RenderTexture Tex_Item(PicNum), ConvertMapX(MapItem(ItemNum).x * PIC_X), ConvertMapY(MapItem(ItemNum).y * PIC_Y), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawMapItem", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawMapResource(ByVal Resource_num As Long)
    Dim Resource_Master As Long
    Dim Resource_State As Long
    Dim Resource_Sprite As Long
    Dim rec As RECT
    Dim x As Long, y As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Make sure it's not out of map
    If MapResource(Resource_num).x > Map.MaxX Then Exit Sub
    If MapResource(Resource_num).y > Map.MaxY Then Exit Sub
    
    ' Get the Resource type
    Resource_Master = Map.Tile(MapResource(Resource_num).x, MapResource(Resource_num).y).Data1
    
    If Resource_Master = 0 Then Exit Sub

    If Resource(Resource_Master).ResourceImage = 0 Then Exit Sub
    
    ' Get the Resource state
    Resource_State = MapResource(Resource_num).ResourceState

    If Resource_State = 0 Then ' Normal
        Resource_Sprite = Resource(Resource_Master).ResourceImage
    ElseIf Resource_State = 1 Then ' Used
        Resource_Sprite = Resource(Resource_Master).ExhaustedImage
    End If
    
    ' Cut down everything if we're editing
    If InMapEditor Then
        Resource_Sprite = Resource(Resource_Master).ExhaustedImage
    End If

    ' Src rect
    With rec
        .Top = 0
        .Bottom = Tex_Resource(Resource_Sprite).Height
        .Left = 0
        .Right = Tex_Resource(Resource_Sprite).Width
    End With

    ' Set base x + y, then the offset due to size
    x = (MapResource(Resource_num).x * PIC_X) - (Tex_Resource(Resource_Sprite).Width / 2) + 16
    y = (MapResource(Resource_num).y * PIC_Y) - Tex_Resource(Resource_Sprite).Height + 32
    
    ' Render it
    Call DrawResource(Resource_Sprite, x, y, rec)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawMapResource", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub DrawResource(ByVal Resource As Long, ByVal dX As Long, dY As Long, rec As RECT)
    Dim x As Long
    Dim y As Long
    Dim Width As Long
    Dim Height As Long
    Dim destRECT As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Resource < 1 Or Resource > NumResources Then Exit Sub

    x = ConvertMapX(dX)
    y = ConvertMapY(dY)
    
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)
    
    RenderTexture Tex_Resource(Resource), x, y, rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawResource", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub DrawBars()
    Dim TmpY As Long, TmpX As Long
    Dim sWidth As Long, sHeight As Long
    Dim sRect As RECT
    Dim i As Long, NpcNum As Long, PartyIndex As Long, BarWidth As Long, MoveSpeed As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Dynamic bar calculations
    sWidth = Tex_Bars.Width
    sHeight = Tex_Bars.Height / 4
    
    ' Render health bars and casting bar
    For i = 1 To MAX_MAP_NPCS
        NpcNum = MapNPC(i).Num
        ' Exists
        If NpcNum > 0 Then
            If Options.NpcVitals = 1 Then
                ' Alive
                If MapNPC(i).Vital(Vitals.HP) < NPC(NpcNum).HP Then
                    ' lock to npc
                    TmpX = MapNPC(i).x * PIC_X + MapNPC(i).xOffset + 16 - (sWidth / 2)
                    TmpY = MapNPC(i).y * PIC_Y + MapNPC(i).yOffset + 35
                    
                    ' Calculate the width to fill
                    BarWidth = ((MapNPC(i).Vital(Vitals.HP) / sWidth) / (NPC(NpcNum).HP / sWidth)) * sWidth
                    
                    ' Draw bar background
                    With sRect
                        .Top = sHeight * 3 ' HP bar background
                        .Left = 0
                        .Right = .Left + sWidth
                        .Bottom = .Top + sHeight
                    End With
                    
                    RenderTexture Tex_Bars, ConvertMapX(TmpX), ConvertMapY(TmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                    
                    ' Draw the bar proper
                    With sRect
                        .Top = 0 ' HP bar
                        .Left = 0
                        .Right = .Left + BarWidth
                        .Bottom = .Top + sHeight
                    End With
                    
                    RenderTexture Tex_Bars, ConvertMapX(TmpX), ConvertMapY(TmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
        
                If MapNPC(i).Vital(Vitals.MP) < NPC(NpcNum).MP Then
                    ' lock to npc
                    TmpX = MapNPC(i).x * PIC_X + MapNPC(i).xOffset + 16 - (sWidth / 2)
                    
                    If MapNPC(i).Vital(Vitals.HP) = NPC(NpcNum).HP Then
                        TmpY = MapNPC(i).y * PIC_Y + MapNPC(i).yOffset + 35
                    Else
                        TmpY = MapNPC(i).y * PIC_Y + MapNPC(i).yOffset + 35 + sHeight
                    End If
                    
                    ' Calculate the width to fill
                    BarWidth = ((MapNPC(i).Vital(Vitals.MP) / sWidth) / (NPC(NpcNum).MP / sWidth)) * sWidth
                    
                    ' Draw bar background
                    With sRect
                        .Top = sHeight * 3 ' MP bar background
                        .Left = 0
                        .Right = .Left + sWidth
                        .Bottom = .Top + sHeight
                    End With
                    
                    RenderTexture Tex_Bars, ConvertMapX(TmpX), ConvertMapY(TmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                    
                    ' Draw the bar proper
                    With sRect
                        .Top = sHeight * 1 ' MP bar
                        .Left = 0
                        .Right = .Left + BarWidth
                        .Bottom = .Top + sHeight
                    End With
                    
                    RenderTexture Tex_Bars, ConvertMapX(TmpX), ConvertMapY(TmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
            End If
            
            ' Check for npc casting time bar
            If MapNPC(i).SpellBuffer > 0 Then
                If MapNPC(i).SpellBufferTimer > timeGetTime - (Spell(MapNPC(i).SpellBuffer).CastTime * 1000) Then
                    ' lock to player
                    TmpX = MapNPC(i).x * PIC_X + MapNPC(i).xOffset + 16 - (sWidth / 2)

                    If Options.NpcVitals = 0 Or (MapNPC(i).Vital(Vitals.HP) = NPC(NpcNum).HP And MapNPC(i).Vital(Vitals.MP) = NPC(NpcNum).MP) Then
                        TmpY = MapNPC(i).y * PIC_Y + MapNPC(i).yOffset + 35
                    Else
                        TmpY = MapNPC(i).y * PIC_Y + MapNPC(i).yOffset + 35 + sHeight
                    End If
                   
                    ' Calculate the width to fill
                    BarWidth = (timeGetTime - MapNPC(i).SpellBufferTimer) / ((Spell(MapNPC(i).SpellBuffer).CastTime * 1000)) * sWidth

                    ' Draw bar background
                    With sRect
                        .Top = sHeight * 3 ' Cooldown bar background
                        .Left = 0
                        .Right = sWidth
                        .Bottom = .Top + sHeight
                    End With
                    
                    RenderTexture Tex_Bars, ConvertMapX(TmpX), ConvertMapY(TmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                    
                    ' Draw the bar proper
                    With sRect
                        .Top = sHeight * 2 ' Cooldown bar
                        .Left = 0
                        .Right = BarWidth
                        .Bottom = .Top + sHeight
                        
                    End With
                    RenderTexture Tex_Bars, ConvertMapX(TmpX), ConvertMapY(TmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
            End If
        End If
    Next
    
    If Options.PlayerVitals = 1 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                ' Draw own health bar
                If GetPlayerVital(i, Vitals.HP) < GetPlayerMaxVital(i, Vitals.HP) Then
                    ' lock to Player
                    TmpX = GetPlayerX(i) * PIC_X + TempPlayer(i).xOffset + 16 - (sWidth / 2)
                    TmpY = GetPlayerY(i) * PIC_X + TempPlayer(i).yOffset + 35
                
                    ' Calculate the width to fill
                    BarWidth = ((GetPlayerVital(i, Vitals.HP) / sWidth) / (GetPlayerMaxVital(i, Vitals.HP) / sWidth)) * sWidth
                    
                    ' Draw bar background
                    With sRect
                        .Top = sHeight * 3 ' HP bar background
                        .Left = 0
                        .Right = .Left + sWidth
                        .Bottom = .Top + sHeight
                    End With
                    
                    RenderTexture Tex_Bars, ConvertMapX(TmpX), ConvertMapY(TmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                   
                    ' Draw the bar proper
                    With sRect
                        .Top = 0 ' HP bar
                        .Left = 0
                        .Right = .Left + BarWidth
                        .Bottom = .Top + sHeight
                    End With
                    
                    RenderTexture Tex_Bars, ConvertMapX(TmpX), ConvertMapY(TmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
                
                ' Draw own mana bar
                If GetPlayerVital(i, Vitals.MP) < GetPlayerMaxVital(i, Vitals.MP) Then
                    ' lock to Player
                    TmpX = GetPlayerX(i) * PIC_X + TempPlayer(i).xOffset + 16 - (sWidth / 2)
                    
                    If GetPlayerVital(i, HP) = GetPlayerMaxVital(i, Vitals.HP) Then
                        TmpY = GetPlayerY(i) * PIC_Y + TempPlayer(i).yOffset + 35
                    Else
                        TmpY = GetPlayerY(i) * PIC_Y + TempPlayer(i).yOffset + 35 + sHeight
                    End If
                   
                    ' Calculate the width to fill
                    BarWidth = ((GetPlayerVital(i, Vitals.MP) / sWidth) / (GetPlayerMaxVital(i, Vitals.MP) / sWidth)) * sWidth
                   
                    ' Draw bar background
                    With sRect
                        .Top = sHeight * 3 ' MP bar background
                        .Left = 0
                        .Right = .Left + sWidth
                        .Bottom = .Top + sHeight
                    End With
                    
                    RenderTexture Tex_Bars, ConvertMapX(TmpX), ConvertMapY(TmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                   
                    ' Draw the bar proper
                    With sRect
                        .Top = sHeight * 1 ' MP bar
                        .Left = 0
                        .Right = .Left + BarWidth
                        .Bottom = .Top + sHeight
                    End With
                    
                    RenderTexture Tex_Bars, ConvertMapX(TmpX), ConvertMapY(TmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
            End If
        Next
    End If
                
    ' Check for player casting time bar
    If SpellBuffer > 0 Then
        If Spell(PlayerSpells(SpellBuffer)).CastTime > 0 Then
            ' lock to player
            TmpX = GetPlayerX(MyIndex) * PIC_X + TempPlayer(MyIndex).xOffset + 16 - (sWidth / 2)
            
            If Options.PlayerVitals = 0 Or (GetPlayerVital(i, HP) = GetPlayerMaxVital(i, Vitals.HP) And GetPlayerVital(i, MP) = GetPlayerMaxVital(i, MP)) Then
                TmpY = GetPlayerY(MyIndex) * PIC_Y + TempPlayer(MyIndex).yOffset + 35
            Else
                TmpY = GetPlayerY(MyIndex) * PIC_Y + TempPlayer(MyIndex).yOffset + 35 + sHeight
            End If
            
            ' Calculate the width to fill
            BarWidth = (timeGetTime - SpellBufferTimer) / ((Spell(PlayerSpells(SpellBuffer)).CastTime * 1000)) * sWidth
            
            ' Draw bar background
            With sRect
                .Top = sHeight * 3 ' Cooldown bar background
                .Left = 0
                .Right = sWidth
                .Bottom = .Top + sHeight
            End With
            
            RenderTexture Tex_Bars, ConvertMapX(TmpX), ConvertMapY(TmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
            
            ' Draw the bar proper
            With sRect
                .Top = sHeight * 2 ' Cooldown bar
                .Left = 0
                .Right = BarWidth
                .Bottom = .Top + sHeight
            End With
            
            RenderTexture Tex_Bars, ConvertMapX(TmpX), ConvertMapY(TmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
        End If
    End If
    
    ' Draw party health bars
    If Party.Num > 0 Then
        For i = 1 To MAX_PARTY_MEMBERS
            PartyIndex = Party.Member(i)
            If (PartyIndex > 0) And (Not PartyIndex = MyIndex) And (GetPlayerMap(PartyIndex) = GetPlayerMap(MyIndex)) Then
                ' player exists
                If GetPlayerVital(PartyIndex, Vitals.HP) > 0 And GetPlayerVital(PartyIndex, Vitals.HP) < GetPlayerMaxVital(PartyIndex, Vitals.HP) Then
                    ' lock to Player
                    TmpX = GetPlayerX(PartyIndex) * PIC_X + TempPlayer(PartyIndex).xOffset + 16 - (sWidth / 2)
                    TmpY = GetPlayerY(PartyIndex) * PIC_X + TempPlayer(PartyIndex).yOffset + 35
                    
                    ' Calculate the width to fill
                    BarWidth = ((GetPlayerVital(PartyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(PartyIndex, Vitals.HP) / sWidth)) * sWidth
                    
                    ' Draw bar background
                    With sRect
                        .Top = sHeight * 1 ' HP bar background
                        .Left = 0
                        .Right = .Left + sWidth
                        .Bottom = .Top + sHeight
                    End With
                    
                    RenderTexture Tex_Bars, ConvertMapX(TmpX), ConvertMapY(TmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                    
                    ' Draw the bar proper
                    With sRect
                        .Top = 0 ' HP bar
                        .Left = 0
                        .Right = .Left + BarWidth
                        .Bottom = .Top + sHeight
                    End With
                    
                    RenderTexture Tex_Bars, ConvertMapX(TmpX), ConvertMapY(TmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
            End If
        Next
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawBars", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawHotbar()
    Dim sRect As RECT, dRect As RECT, i As Long, Num As String, n As Long, destRECT As D3DRECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    For i = 1 To MAX_HOTBAR
        With dRect
            .Top = HotbarTop
            .Left = HotbarLeft + ((HotbarOffsetX + 32) * (((i - 1) Mod MAX_HOTBAR)))
            .Bottom = .Top + 32
            .Right = .Left + 32
        End With
        
        With destRECT
            .Y1 = HotbarTop
            .X1 = HotbarLeft + ((HotbarOffsetX + 32) * (((i - 1) Mod MAX_HOTBAR)))
            .Y2 = .Y1 + 32
            .X2 = .X1 + 32
        End With
        
        With sRect
            .Top = 0
            .Left = 32
            .Bottom = 32
            .Right = 64
        End With
    
        Select Case Hotbar(i).SType
            Case 1 ' Inventory
                If Len(Item(Hotbar(i).Slot).name) > 0 Then
                    If Item(Hotbar(i).Slot).Pic > 0 Then
                        If Item(Hotbar(i).Slot).Pic <= NumItems Then
                            Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 255), 1#, 0
                            Direct3D_Device.BeginScene
                            RenderTextureByRects Tex_Item(Item(Hotbar(i).Slot).Pic), sRect, dRect
                            Direct3D_Device.EndScene
                            Direct3D_Device.Present destRECT, destRECT, frmMain.picHotbar.hwnd, ByVal (0)
                        End If
                    End If
                End If
            Case 2 ' Spell
                If Len(Spell(Hotbar(i).Slot).name) > 0 Then
                    If Spell(Hotbar(i).Slot).Icon > 0 Then
                        With sRect
                            .Top = 0
                            .Left = 0
                            .Bottom = 32
                            .Right = 32
                        End With
                        
                        ' Check for cooldown
                        For n = 1 To MAX_PLAYER_SPELLS
                            If PlayerSpells(n) = Hotbar(i).Slot Then
                                ' has spell
                                If Not SpellCD(n) = 0 Then
                                    sRect.Left = 32
                                    sRect.Right = 64
                                    Exit For
                                End If
                            End If
                        Next
                        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 255), 1#, 0
                        Direct3D_Device.BeginScene
                        RenderTextureByRects Tex_SpellIcon(Spell(Hotbar(i).Slot).Icon), sRect, dRect
                        Direct3D_Device.EndScene
                        Direct3D_Device.Present destRECT, destRECT, frmMain.picHotbar.hwnd, ByVal (0)
                    End If
                End If
        End Select

        ' Render the letters
        If Options.WASD = 1 Then
            If i = 10 Then
                Num = " 0"
            ElseIf i = 11 Then
                Num = " -"
            ElseIf i = 12 Then
                Num = " +"
            Else
                Num = " " & Trim$(i)
            End If
        Else
            Num = " F" & Trim$(i)
        End If
        RenderText Font_Default, Num, dRect.Left + 2, dRect.Top + 16, White
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawHotbar", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawPlayer(ByVal Index As Long)
    Dim Anim As Byte, i As Long, x As Long, y As Long
    Dim Sprite As Long, spritetop As Long
    Dim rec As RECT
    Dim AttackSpeed As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Sprite = GetPlayerSprite(Index)

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    ' Speed from weapon
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        AttackSpeed = Item(GetPlayerEquipment(Index, Weapon)).WeaponSpeed
    Else
        AttackSpeed = 1000
    End If

    ' Reset frame
    If TempPlayer(Index).Moving > 0 Then
        Anim = TempPlayer(Index).Step
    Else
        Anim = 0
    End If
    
    ' If the sprite is constantly animated, make it animate
    If Not IsConstAnimated(GetPlayerSprite(Index)) Then
        ' Check for attacking animation
        If TempPlayer(Index).AttackTimer + (AttackSpeed / 2) > timeGetTime Then
            If TempPlayer(Index).Attacking = 1 Then
                Anim = 3
            End If
        Else
            ' If not attacking, walk normally
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    If (TempPlayer(Index).yOffset > 8) Then Anim = TempPlayer(Index).Step
                Case DIR_DOWN
                    If (TempPlayer(Index).yOffset < -8) Then Anim = TempPlayer(Index).Step
                Case DIR_LEFT
                    If (TempPlayer(Index).xOffset > 8) Then Anim = TempPlayer(Index).Step
                Case DIR_RIGHT
                    If (TempPlayer(Index).xOffset < -8) Then Anim = TempPlayer(Index).Step
            End Select
        End If
    
        ' Check to see if we want to stop making him attack
        With TempPlayer(Index)
            If .AttackTimer + AttackSpeed < timeGetTime Then
                .Attacking = 0
                .AttackTimer = 0
            End If
        End With
    Else
        If TempPlayer(Index).AnimTimer + 100 <= timeGetTime Then
            TempPlayer(Index).Anim = TempPlayer(Index).Anim + 1
            If TempPlayer(Index).Anim >= 4 Then TempPlayer(Index).Anim = 0
            TempPlayer(Index).AnimTimer = timeGetTime
        End If
        Anim = TempPlayer(Index).Anim
    End If

    ' Set the left
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
    End Select

    With rec
        .Top = spritetop * (Tex_Character(Sprite).Height / 4)
        .Bottom = .Top + (Tex_Character(Sprite).Height / 4)
        .Left = Anim * (Tex_Character(Sprite).Width / 4)
        .Right = .Left + (Tex_Character(Sprite).Width / 4)
    End With

    ' Calculate the X
    x = GetPlayerX(Index) * PIC_X + TempPlayer(Index).xOffset - ((Tex_Character(Sprite).Width / 4 - 32) / 2)

    ' Is the player's height more than 32?
    If (Tex_Character(Sprite).Height) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        y = GetPlayerY(Index) * PIC_Y + TempPlayer(Index).yOffset - ((Tex_Character(Sprite).Height / 4) - 32)
    Else
        ' Proceed as normal
        y = GetPlayerY(Index) * PIC_Y + TempPlayer(Index).yOffset
    End If

    ' Render the actual sprite
    Call DrawSprite(Sprite, x, y, rec)
    
    ' Check for paperdolling
    For i = 1 To UBound(PaperdollOrder)
        If GetPlayerEquipment(Index, PaperdollOrder(i)) > 0 Then
            If Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll > 0 Then
                Call DrawPaperdoll(x, y, Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll, Anim, spritetop)
            End If
        End If
    Next
    Exit Sub

' Error handler
errorhandler:
    HandleError "DrawPlayer", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawNpc(ByVal MapNPCNum As Long)
    Dim Anim As Byte, i As Long, x As Long, y As Long, Sprite As Long, spritetop As Long
    Dim rec As RECT
    Dim AttackSpeed As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If MapNPC(MapNPCNum).Num = 0 Then Exit Sub ' No npc set
    
    Sprite = NPC(MapNPC(MapNPCNum).Num).Sprite

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    AttackSpeed = 1000

    ' Reset frame
    Anim = 0
    
    If Not IsConstAnimated(NPC(MapNPC(MapNPCNum).Num).Sprite) Then
        ' Check for attacking animation
        If MapNPC(MapNPCNum).AttackTimer + (AttackSpeed / 2) > timeGetTime Then
            If MapNPC(MapNPCNum).Attacking = 1 Then
                Anim = 3
            End If
        Else
            ' If not attacking, walk normally
            Select Case MapNPC(MapNPCNum).Dir
                Case DIR_UP
                    If (MapNPC(MapNPCNum).yOffset > 8) Then Anim = MapNPC(MapNPCNum).Step
                Case DIR_DOWN
                    If (MapNPC(MapNPCNum).yOffset < -8) Then Anim = MapNPC(MapNPCNum).Step
                Case DIR_LEFT
                    If (MapNPC(MapNPCNum).xOffset > 8) Then Anim = MapNPC(MapNPCNum).Step
                Case DIR_RIGHT
                    If (MapNPC(MapNPCNum).xOffset < -8) Then Anim = MapNPC(MapNPCNum).Step
            End Select
        End If
    Else
        With MapNPC(MapNPCNum)
            If .AnimTimer + 100 <= timeGetTime Then
                .Anim = .Anim + 1
                If .Anim >= 4 Then .Anim = 0
                .AnimTimer = timeGetTime
            End If
            Anim = .Anim
        End With
    End If

    ' Check to see if we want to stop making him attack
    With MapNPC(MapNPCNum)
        If .AttackTimer + AttackSpeed < timeGetTime Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the left
    Select Case MapNPC(MapNPCNum).Dir
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
    End Select

    With rec
        .Top = (Tex_Character(Sprite).Height / 4) * spritetop
        .Bottom = .Top + Tex_Character(Sprite).Height / 4
        .Left = Anim * (Tex_Character(Sprite).Width / 4)
        .Right = .Left + (Tex_Character(Sprite).Width / 4)
    End With

    ' Calculate the X
    x = MapNPC(MapNPCNum).x * PIC_X + MapNPC(MapNPCNum).xOffset - ((Tex_Character(Sprite).Width / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (Tex_Character(Sprite).Height / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        y = MapNPC(MapNPCNum).y * PIC_Y + MapNPC(MapNPCNum).yOffset - ((Tex_Character(Sprite).Height / 4) - 32)
    Else
        ' Proceed as normal
        y = MapNPC(MapNPCNum).y * PIC_Y + MapNPC(MapNPCNum).yOffset
    End If

    Call DrawSprite(Sprite, x, y, rec)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawNpc", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawPaperdoll(ByVal X2 As Long, ByVal Y2 As Long, ByVal Sprite As Long, ByVal Anim As Long, ByVal spritetop As Long)
    Dim rec As RECT
    Dim x As Long, y As Long
    Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Sprite < 1 Or Sprite > NumPaperdolls Then Exit Sub
    
    With rec
        .Top = spritetop * (Tex_Paperdoll(Sprite).Height / 4)
        .Bottom = .Top + (Tex_Paperdoll(Sprite).Height / 4)
        .Left = Anim * (Tex_Paperdoll(Sprite).Width / 4)
        .Right = .Left + (Tex_Paperdoll(Sprite).Width / 4)
    End With
    
    ' Clipping
    x = ConvertMapX(X2)
    y = ConvertMapY(Y2)
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)

    ' Clip to screen
    If y < 0 Then
        With rec
            .Top = .Top - y
        End With
        y = 0
    End If

    If x < 0 Then
        With rec
            .Left = .Left - x
        End With
        x = 0
    End If
    
    RenderTexture Tex_Paperdoll(Sprite), x, y, rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawPaperdoll", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub DrawSprite(ByVal Sprite As Long, ByVal X2 As Long, Y2 As Long, rec As RECT)
    Dim x As Long
    Dim y As Long
    Dim Width As Long
    Dim Height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    
    x = ConvertMapX(X2)
    y = ConvertMapY(Y2)
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)
    
    RenderTexture Tex_Character(Sprite), x, y, rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawSprite", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub DrawAnimatedItems()
    Dim i As Long
    Dim ItemNum As Long, ItemPic As Long, Color As Long
    Dim x As Long, y As Long
    Dim MaxFrames As Byte
    Dim Amount As Long
    Dim rec As RECT, rec_pos As RECT
    Dim TmpItem As Long, AmountModifier As Long
    Dim NoRender(1 To MAX_INV) As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    ' Check for map animation changes
    For i = 1 To MAX_MAP_ITEMS
        If MapItem(i).Num > 0 Then
            ItemPic = Item(MapItem(i).Num).Pic

            If ItemPic < 1 Or ItemPic > NumItems Then Exit Sub
            MaxFrames = (Tex_Item(ItemPic).Width / 2) / 32 ' Work out how many frames there are. /2 because of inventory icons as well as ingame

            If MapItem(i).Frame < MaxFrames - 1 Then
                MapItem(i).Frame = MapItem(i).Frame + 1
            Else
                MapItem(i).Frame = 1
            End If
        End If
    Next
    
    For i = 1 To MAX_INV
        ItemNum = GetPlayerInvItemNum(MyIndex, i)
        
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ItemPic = Item(ItemNum).Pic
            AmountModifier = 0
            NoRender(i) = 0
            
            ' Exit out if we're offering item in a trade.
            If InTrade > 0 Then
                For x = 1 To MAX_INV
                    TmpItem = GetPlayerInvItemNum(MyIndex, TradeYourOffer(x).Num)
                    If TradeYourOffer(x).Num = i Then
                        ' Check if currency
                        If Not Item(TmpItem).Stackable = 1 Then
                            ' Normal item don't render
                            NoRender(i) = 1
                        Else
                            ' If amount = all currency, remove from inventory
                            If TradeYourOffer(x).Value = GetPlayerInvItemValue(MyIndex, i) Then
                                NoRender(i) = 1
                            Else
                                ' Not all, change modifier to show change in currency count
                                AmountModifier = TradeYourOffer(x).Value
                            End If
                        End If
                    End If
                Next
            End If
                
            If NoRender(i) = 0 Then
                If ItemPic > 0 And ItemPic <= NumItems Then
                    If Tex_Item(ItemPic).Width > 64 Then
                        MaxFrames = (Tex_Item(ItemPic).Width / 2) / 32 ' Work out how many frames there are. /2 because of inventory icons as well as ingame
    
                        If InvItemFrame(i) < MaxFrames - 1 Then
                            InvItemFrame(i) = InvItemFrame(i) + 1
                        Else
                            InvItemFrame(i) = 1
                        End If
    
                        With rec
                            .Top = 0
                            .Bottom = 32
                            .Left = (Tex_Item(ItemPic).Width / 2) + (InvItemFrame(i) * 32) ' Middle to get the start of inv gfx, then +32 for each frame
                            .Right = .Left + 32
                        End With
    
                        With rec_pos
                            .Top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                            .Bottom = .Top + PIC_Y
                            .Left = InvLeft + ((InvOffsetX + PIC_X) * (((i - 1) Mod InvColumns)))
                            .Right = .Left + PIC_X
                        End With

                        ' We'll now re-Draw the item, and place the currency value over it again :P
                        RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos
    
                        ' If item is a stack - draw the amount you have
                        If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                            y = rec_pos.Top + 22
                            x = rec_pos.Left - 4
                            Amount = GetPlayerInvItemValue(MyIndex, i) - AmountModifier
                            
                            ' Draw currency but with k, m, b etc. using a convertion function
                            If Amount < 1000000 Then
                                Color = White
                            ElseIf Amount > 1000000 And Amount < 10000000 Then
                                Color = Yellow
                            ElseIf Amount > 10000000 Then
                                Color = BrightGreen
                            End If
                            
                            ' Draw currency but with k, m, b etc. using a convertion function
                            RenderText Font_Default, ConvertCurrency(Amount), x, y, Color
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    If InBank Then
        For i = 1 To MAX_BANK
            ItemNum = GetBankItemNum(i)
            
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
                ItemPic = Item(ItemNum).Pic
    
                If ItemPic > 0 And ItemPic <= NumItems Then
                    If Tex_Item(ItemPic).Width > 64 Then
                        MaxFrames = (Tex_Item(ItemPic).Width / 2) / 32 ' Work out how many frames there are. /2 because of bankentory icons as well as ingame
    
                        If BankItemFrame(i) < MaxFrames - 1 Then
                            BankItemFrame(i) = BankItemFrame(i) + 1
                        Else
                            BankItemFrame(i) = 1
                        End If
    
                        With rec
                            .Top = 0
                            .Bottom = 32
                            .Left = (Tex_Item(ItemPic).Width / 2) + (BankItemFrame(i) * 32) ' Middle to get the start of Bank gfx, then +32 for each frame
                            .Right = .Left + 32
                        End With
    
                        With rec_pos
                            .Top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                            .Bottom = .Top + PIC_Y
                            .Left = BankLeft + ((BankOffsetX + PIC_X) * (((i - 1) Mod BankColumns)))
                            .Right = .Left + PIC_X
                        End With
    
                        ' We'll now re-Draw the item, and place the currency value over it again :P
                        RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos
    
                        ' If item is a stack - draw the amount you have
                        If GetBankItemValue(i) > 1 Then
                            y = rec_pos.Top + 22
                            x = rec_pos.Left - 4
                            Amount = GetBankItemValue(i)
                            
                            ' Draw currency but with k, m, b etc. using a convertion function
                            If Amount < 1000000 Then
                                Color = White
                            ElseIf Amount > 1000000 And Amount < 10000000 Then
                                Color = Yellow
                            ElseIf Amount > 10000000 Then
                                Color = BrightGreen
                            End If
                            
                            ' Draw currency but with k, m, b etc. using a convertion function
                            RenderText Font_Default, ConvertCurrency(Amount), x, y, Color
                        End If
                    End If
                End If
            End If
        Next
    End If
    
    If InShop > 0 Then
        For i = 1 To MAX_TRADES
            ItemNum = Shop(InShop).TradeItem(i).Item
            
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
                ItemPic = Item(ItemNum).Pic
    
                If ItemPic > 0 And ItemPic <= NumItems Then
                    If Tex_Item(ItemPic).Width > 64 Then
                        MaxFrames = (Tex_Item(ItemPic).Width / 2) / 32 ' Work out how many frames there are. /2 because of shopentory icons as well as ingame
    
                        If ShopItemFrame(i) < MaxFrames - 1 Then
                            ShopItemFrame(i) = ShopItemFrame(i) + 1
                        Else
                            ShopItemFrame(i) = 1
                        End If
    
                        With rec
                            .Top = 0
                            .Bottom = 32
                            .Left = (Tex_Item(ItemPic).Width / 2) + (ShopItemFrame(i) * 32) ' Middle to get the start of shop gfx, then +32 for each frame
                            .Right = .Left + 32
                        End With
    
                        With rec_pos
                            .Top = ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
                            .Bottom = .Top + PIC_Y
                            .Left = ShopLeft + ((ShopOffsetX + PIC_X) * (((i - 1) Mod ShopColumns)))
                            .Right = .Left + PIC_X
                        End With
                        
                        ' If item is a stack - draw the amount you have
                        If Shop(InShop).TradeItem(i).ItemValue > 1 Then
                            y = rec_pos.Top + 22
                            x = rec_pos.Left - 4
                            Amount = Shop(InShop).TradeItem(i).ItemValue
                            
                            ' Draw currency but with k, m, b etc. using a convertion function
                            If Amount < 1000000 Then
                                Color = White
                            ElseIf Amount > 1000000 And Amount < 10000000 Then
                                Color = Yellow
                            ElseIf Amount > 10000000 Then
                                Color = BrightGreen
                            End If
                            
                            ' Draw currency but with k, m, b etc. using a convertion function
                            RenderText Font_Default, ConvertCurrency(Amount), x, y, Color
                        End If
    
                        ' We'll now re-Draw the item, and place the currency value over it again :P
                        RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos
                    End If
                End If
            End If
        Next
    End If
    
    If frmMain.picTrade.Visible = True Then
        For i = 1 To MAX_INV
            ItemNum = TradeTheirOffer(i).Num
            
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
                ItemPic = Item(ItemNum).Pic
    
                If ItemPic > 0 And ItemPic <= NumItems Then
                    If Tex_Item(ItemPic).Width > 64 Then
                        MaxFrames = (Tex_Item(ItemPic).Width / 2) / 32 ' Work out how many frames there are. /2 because of TheirTrade icons as well as ingame
    
                        If InvItemFrame(i) < MaxFrames - 1 Then
                            InvItemFrame(i) = InvItemFrame(i) + 1
                        Else
                            InvItemFrame(i) = 1
                        End If
    
                        With rec
                            .Top = 0
                            .Bottom = 32
                            .Left = (Tex_Item(ItemPic).Width / 2) + (InvItemFrame(i) * 32) ' Middle to get the start of inv gfx, then +32 for each frame
                            .Right = .Left + 32
                        End With
    
                        With rec_pos
                            .Top = InvTop - 12 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                            .Bottom = .Top + PIC_Y
                            .Left = InvLeft + ((InvOffsetX + PIC_X) * (((i - 1) Mod InvColumns)))
                            .Right = .Left + PIC_X
                        End With

                        ' We'll now re-Draw the item, and place the currency value over it again :P
                        RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos
                    
                        ' If item is a stack - draw the amount you have
                        If TradeTheirOffer(i).Value > 1 Then
                            y = rec_pos.Top + 22
                            x = rec_pos.Left - 4
                            Amount = TradeTheirOffer(i).Value
                            
                            ' Draw currency but with k, m, b etc. using a convertion function
                            If Amount < 1000000 Then
                                Color = White
                            ElseIf Amount > 1000000 And Amount < 10000000 Then
                                Color = Yellow
                            ElseIf Amount > 10000000 Then
                                Color = BrightGreen
                            End If
                            
                            ' Draw currency but with k, m, b etc. using a convertion function
                            RenderText Font_Default, ConvertCurrency(Amount), x, y, Color
                        End If
                    End If
                End If
            End If
        Next
        
         For i = 1 To MAX_INV
            ItemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).Num)
            
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
                ItemPic = Item(ItemNum).Pic
    
                If ItemPic > 0 And ItemPic <= NumItems Then
                    If Tex_Item(ItemPic).Width > 64 Then
                        MaxFrames = (Tex_Item(ItemPic).Width / 2) / 32 ' Work out how many frames there are. /2 because of YourTrade icons as well as ingame
    
                        If InvItemFrame(i) < MaxFrames - 1 Then
                            InvItemFrame(i) = InvItemFrame(i) + 1
                        Else
                            InvItemFrame(i) = 1
                        End If
    
                        With rec
                            .Top = 0
                            .Bottom = 32
                            .Left = (Tex_Item(ItemPic).Width / 2) + (InvItemFrame(i) * 32) ' Middle to get the start of inv gfx, then +32 for each frame
                            .Right = .Left + 32
                        End With
    
                        With rec_pos
                            .Top = InvTop - 12 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                            .Bottom = .Top + PIC_Y
                            .Left = InvLeft + ((InvOffsetX + PIC_X) * (((i - 1) Mod InvColumns)))
                            .Right = .Left + PIC_X
                        End With

                        ' We'll now re-Draw the item, and place the currency value over it again :P
                        RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos
                        
                        ' If item is a stack - draw the amount you have
                        If TradeYourOffer(i).Value > 1 Then
                            y = rec_pos.Top + 22
                            x = rec_pos.Left - 4
                            Amount = TradeYourOffer(i).Value
                            
                            ' Draw currency but with k, m, b etc. using a convertion function
                            If Amount < 1000000 Then
                                Color = White
                            ElseIf Amount > 1000000 And Amount < 10000000 Then
                                Color = Yellow
                            ElseIf Amount > 10000000 Then
                                Color = BrightGreen
                            End If
                            
                            ' Draw currency but with k, m, b etc. using a convertion function
                            RenderText Font_Default, ConvertCurrency(Amount), x, y, Color
                        End If
                    End If
                End If
            End If
        Next
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawAnimatedItems", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub DrawPlayerCharFace()
    Dim rec As RECT, rec_pos As RECT, FaceNum As Long, srcRect As D3DRECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If NumFaces = 0 Then Exit Sub
    
    FaceNum = Player(MyIndex).Face
    
    If FaceNum <= 0 Or FaceNum > NumFaces Then Exit Sub

    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene

    With rec
        .Top = 0
        .Bottom = Tex_Face(FaceNum).Height
        .Left = 0
        .Right = Tex_Face(FaceNum).Width
    End With

    With rec_pos
        .Top = 0
        .Bottom = 100
        .Left = 0
        .Right = 100
    End With

    RenderTextureByRects Tex_Face(FaceNum), rec, rec_pos
    
    With srcRect
        .X1 = 0
        .X2 = frmMain.picFace.Width
        .Y1 = 0
        .Y2 = frmMain.picFace.Height
    End With
    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, srcRect, frmMain.picFace.hwnd, ByVal (0)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawPlayerCharFace", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub DrawInventory()
    Dim i As Long, x As Long, y As Long, ItemNum As Long, ItemPic As Long
    Dim Amount As Long
    Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT, destRECT As D3DRECT
    Dim Color As Long
    Dim TmpItem As Long, AmountModifier As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene

    For i = 1 To MAX_INV
        ItemNum = GetPlayerInvItemNum(MyIndex, i)

        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ItemPic = Item(ItemNum).Pic
            AmountModifier = 0
            
            ' Exit out if we're offering item in a trade.
            If InTrade > 0 Then
                For x = 1 To MAX_INV
                    TmpItem = GetPlayerInvItemNum(MyIndex, TradeYourOffer(x).Num)
                    If TradeYourOffer(x).Num = i Then
                        ' Check if currency
                        If Not Item(TmpItem).Stackable = 1 Then
                            ' Normal item, exit out
                            GoTo NextLoop
                        Else
                            ' If amount = all currency, remove from inventory
                            If TradeYourOffer(x).Value = GetPlayerInvItemValue(MyIndex, i) Then
                                GoTo NextLoop
                            Else
                                ' Not all, change modifier to show change in currency count
                                AmountModifier = TradeYourOffer(x).Value
                            End If
                        End If
                    End If
                Next
            End If

            If ItemPic > 0 And ItemPic <= NumItems Then
                If Tex_Item(ItemPic).Width <= 64 Then ' More than 1 frame is handled by anim sub
                     With rec
                        .Top = 0
                        .Bottom = 32
                        .Left = 32
                        .Right = 64
                    End With

                    With rec_pos
                        .Top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                        .Bottom = .Top + PIC_Y
                        .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                        .Right = .Left + PIC_X
                    End With

                    RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos

                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                        y = rec_pos.Top + 22
                        x = rec_pos.Left - 4
                        Amount = GetPlayerInvItemValue(MyIndex, i) - AmountModifier
                        
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If Amount < 1000000 Then
                            Color = White
                        ElseIf Amount > 1000000 And Amount < 10000000 Then
                            Color = Yellow
                        ElseIf Amount > 10000000 Then
                            Color = BrightGreen
                        End If
                        
                        RenderText Font_Default, Format$(ConvertCurrency(str(Amount)), "#,###,###,###"), x, y, Color
                    End If
                End If
            End If
        End If
        
NextLoop:
    Next
    
    With rec
        .Top = 0
        .Bottom = Tex_Base.Height
        .Left = 0
        .Right = Tex_Base.Width
    End With
    
    With rec_pos
        .Top = 0
        .Bottom = frmMain.picInventory.Height
        .Left = 0
        .Right = frmMain.picInventory.Width
    End With

    RenderTextureByRects Tex_Base, rec, rec_pos
    
    With srcRect
        .X1 = 0
        .X2 = frmMain.picInventory.Width
        .Y1 = 0
        .Y2 = frmMain.picInventory.Height
    End With
    
    With destRECT
        .X1 = 0
        .X2 = frmMain.picInventory.Width
        .Y1 = 0
        .Y2 = frmMain.picInventory.Height
    End With
    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, destRECT, frmMain.picInventory.hwnd, ByVal (0)
    
    ' Update animated items
    DrawAnimatedItems
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawInventory", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub DrawTrade()
    Dim i As Long, x As Long, y As Long, ItemNum As Long, ItemPic As Long
    Dim Amount As Long
    Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT, destRECT As D3DRECT
    Dim Color As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    For i = 1 To MAX_INV
        ' Draw your own offer
        ItemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).Num)

        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ItemPic = Item(ItemNum).Pic

            If ItemPic > 0 And ItemPic <= NumItems Then
                If Tex_Item(ItemPic).Width <= 64 Then
                    With rec
                        .Top = 0
                        .Bottom = 32
                        .Left = 32
                        .Right = 64
                    End With
    
                    With rec_pos
                        .Top = InvTop - 12 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                        .Bottom = .Top + PIC_Y
                        .Left = InvLeft + ((InvOffsetX + PIC_X) * (((i - 1) Mod InvColumns)))
                        .Right = .Left + PIC_X
                    End With
    
                    RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos
    
                    ' If item is a stack - draw the amount you have
                    If TradeYourOffer(i).Value > 1 Then
                        y = rec_pos.Top + 22
                        x = rec_pos.Left - 4
                        Amount = TradeYourOffer(i).Value
                        
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If Amount < 1000000 Then
                            Color = White
                        ElseIf Amount > 1000000 And Amount < 10000000 Then
                            Color = Yellow
                        ElseIf Amount > 10000000 Then
                            Color = BrightGreen
                        End If
                        
                        RenderText Font_Default, Format$(ConvertCurrency(str(Amount)), "#,###,###,###"), x, y, Color
                    End If
                End If
            End If
        End If
        
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRECT, frmMain.picYourTrade.hwnd, ByVal (0)
        
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
            
        ' Draw their offer
        ItemNum = TradeTheirOffer(i).Num

        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ItemPic = Item(ItemNum).Pic

            If ItemPic > 0 And ItemPic <= NumItems Then
                If Tex_Item(ItemPic).Width <= 64 Then
                    With rec
                        .Top = 0
                        .Bottom = 32
                        .Left = 32
                        .Right = 64
                    End With
    
                    With rec_pos
                        .Top = InvTop - 12 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                        .Bottom = .Top + PIC_Y
                        .Left = InvLeft + ((InvOffsetX + PIC_X) * (((i - 1) Mod InvColumns)))
                        .Right = .Left + PIC_X
                    End With
    
                    RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos
    
                    ' If item is a stack - draw the amount you have
                    If TradeTheirOffer(i).Value > 1 Then
                        y = rec_pos.Top + 22
                        x = rec_pos.Left - 4
                        Amount = TradeTheirOffer(i).Value
                        
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If Amount < 1000000 Then
                            Color = White
                        ElseIf Amount > 1000000 And Amount < 10000000 Then
                            Color = Yellow
                        ElseIf Amount > 10000000 Then
                            Color = BrightGreen
                        End If
                        
                        RenderText Font_Default, Format$(ConvertCurrency(str(Amount)), "#,###,###,###"), x, y, Color
                    End If
                End If
            End If
        End If
    Next
    
    With srcRect
        .X1 = 0
        .X2 = .X1 + 193
        .Y1 = 0
        .Y2 = .Y1 + 246
    End With
                    
    With destRECT
        .X1 = 0
        .X2 = .X1 + 193
        .Y1 = 0
        .Y2 = 246 + .Y1
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, destRECT, frmMain.picTheirTrade.hwnd, ByVal (0)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawTrade", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub DrawPlayerSpells()
    Dim i As Long, x As Long, y As Long, SpellNum As Long, SpellIcon As Long, srcRect As D3DRECT, destRECT As D3DRECT
    Dim Amount As String
    Dim rec As RECT, rec_pos As RECT
    Dim Color As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    With rec
        .Top = 0
        .Bottom = Tex_Base.Height
        .Left = 0
        .Right = Tex_Base.Width
    End With
    
    With rec_pos
        .Top = 0
        .Bottom = frmMain.picSpells.Height
        .Left = 0
        .Right = frmMain.picSpells.Width
    End With

    RenderTextureByRects Tex_Base, rec, rec_pos

    For i = 1 To MAX_PLAYER_SPELLS
        SpellNum = PlayerSpells(i)

        If SpellNum > 0 And SpellNum <= MAX_SPELLS Then
            SpellIcon = Spell(SpellNum).Icon

            If SpellIcon > 0 And SpellIcon <= NumSpellIcons Then
                With rec
                    .Top = 0
                    .Bottom = 32
                    .Left = 0
                    .Right = 32
                End With
                
                If Not SpellCD(i) = 0 Then
                    rec.Left = 32
                    rec.Right = 64
                End If

                With rec_pos
                    .Top = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                    .Bottom = .Top + PIC_Y
                    .Left = SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                    .Right = .Left + PIC_X
                End With

                RenderTextureByRects Tex_SpellIcon(SpellIcon), rec, rec_pos
            End If
        End If
    Next
    
    With srcRect
        .X1 = 0
        .X2 = frmMain.picSpells.Width
        .Y1 = 0
        .Y2 = frmMain.picSpells.Height
    End With
    
    With destRECT
        .X1 = 0
        .X2 = frmMain.picSpells.Width
        .Y1 = 0
        .Y2 = frmMain.picSpells.Height
    End With
    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, destRECT, frmMain.picSpells.hwnd, ByVal (0)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawPlayerSpells", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub DrawShop()
    Dim i As Long, x As Long, y As Long, ItemNum As Long, ItemPic As Long, srcRect As D3DRECT, destRECT As D3DRECT
    Dim Amount As String
    Dim rec As RECT, rec_pos As RECT
    Dim Color As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    For i = 1 To MAX_TRADES
        ItemNum = Shop(InShop).TradeItem(i).Item
        
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ItemPic = Item(ItemNum).Pic

            If ItemPic > 0 And ItemPic <= NumItems Then
                If Tex_Item(ItemPic).Width <= 64 Then
                
                    With rec
                        .Top = 0
                        .Bottom = 32
                        .Left = 32
                        .Right = 64
                    End With
                    
                    With rec_pos
                        .Top = ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
                        .Bottom = .Top + PIC_Y
                        .Left = ShopLeft + ((ShopOffsetX + PIC_X) * (((i - 1) Mod ShopColumns)))
                        .Right = .Left + PIC_X
                    End With

                    RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos
                    
                    ' If item is a stack - draw the amount you have
                    If Shop(InShop).TradeItem(i).ItemValue > 1 Then
                        y = rec_pos.Top + 22
                        x = rec_pos.Left - 4
                        Amount = Shop(InShop).TradeItem(i).ItemValue
                        
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If Amount < 1000000 Then
                            Color = White
                        ElseIf Amount > 1000000 And Amount < 10000000 Then
                            Color = Yellow
                        ElseIf Amount > 10000000 Then
                            Color = BrightGreen
                        End If
                        
                        RenderText Font_Default, ConvertCurrency(Amount), x, y, Color
                    End If
                End If
            End If
        End If
    Next
    
    With srcRect
        .X1 = ShopLeft
        .X2 = .X1 + 192
        .Y1 = ShopTop
        .Y2 = .Y1 + 211
    End With
                
    With destRECT
        .X1 = ShopLeft
        .X2 = .X1 + 192
        .Y1 = ShopTop
        .Y2 = 211 + .Y1
    End With
                
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, destRECT, frmMain.picShopItems.hwnd, ByVal (0)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawShop", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawDraggedItem(ByVal x As Long, ByVal y As Long, Optional ByVal IsHotbarSlot As Boolean = False)
    Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT, destRECT As D3DRECT
    Dim ItemNum As Long, ItemPic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If IsHotbarSlot Then
        ItemNum = Hotbar(DragHotbarSlot).Slot
    Else
        ItemNum = GetPlayerInvItemNum(MyIndex, DragInvSlot)
    End If

    If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
    
        ItemPic = Item(ItemNum).Pic
        
        If ItemPic = 0 Then Exit Sub
        
        With rec
            .Top = 0
            .Bottom = .Top + PIC_Y
            .Left = Tex_Item(ItemPic).Width / 2
            .Right = .Left + PIC_X
        End With

        With rec_pos
            .Top = 2
            .Bottom = .Top + PIC_Y
            .Left = 2
            .Right = .Left + PIC_X
        End With

        RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos

        With frmMain.picTempInv
            .Top = y
            .Left = x
            .Visible = True
            .ZOrder (0)
        End With
        
        With srcRect
            .X1 = 0
            .X2 = 32
            .Y1 = 0
            .Y2 = 32
        End With
        
        With destRECT
            .X1 = 2
            .Y1 = 2
            .Y2 = .Y1 + 32
            .X2 = .X1 + 32
        End With
        
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRECT, frmMain.picTempInv.hwnd, ByVal (0)
    End If
    Exit Sub

' Error handler
errorhandler:
    HandleError "DrawDraggedItem", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawDraggedSpell(ByVal x As Long, ByVal y As Long, Optional ByVal IsHotbarSlot As Boolean = False)
    Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT, destRECT As D3DRECT
    Dim SpellNum As Long, SpellPic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If IsHotbarSlot Then
        SpellNum = Hotbar(DragHotbarSlot).Slot
    Else
        SpellNum = PlayerSpells(DragSpellSlot)
    End If
    
     If SpellNum > 0 And SpellNum <= MAX_SPELLS Then
        SpellPic = Spell(SpellNum).Icon
        
        If SpellPic = 0 Then Exit Sub
        
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene

        With rec
            .Top = 0
            .Bottom = .Top + PIC_Y
            .Left = 0
            .Right = .Left + PIC_X
        End With
        
        If IsHotbarSlot = False Then
            If SpellCD(DragSpellSlot) Then
                With rec
                    .Top = 0
                    .Bottom = .Top + PIC_Y
                    .Left = Tex_SpellIcon(SpellPic).Width / 2
                    .Right = .Left + PIC_X
                End With
            End If
        Else
            If SpellCD(DragHotbarSlot) Then
                With rec
                    .Top = 0
                    .Bottom = .Top + PIC_Y
                    .Left = Tex_SpellIcon(SpellPic).Width / 2
                    .Right = .Left + PIC_X
                End With
            End If
        End If

        With rec_pos
            .Top = 2
            .Bottom = .Top + PIC_Y
            .Left = 2
            .Right = .Left + PIC_X
        End With

        RenderTextureByRects Tex_SpellIcon(SpellPic), rec, rec_pos

        With frmMain.picTempSpell
            .Top = y
            .Left = x
            .Visible = True
            .ZOrder (0)
        End With
        
        With srcRect
            .X1 = 0
            .X2 = 32
            .Y1 = 0
            .Y2 = 32
        End With
        With destRECT
            .X1 = 2
            .Y1 = 2
            .Y2 = .Y1 + 32
            .X2 = .X1 + 32
        End With
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRECT, frmMain.picTempSpell.hwnd, ByVal (0)
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawDraggedSpell", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawItemDesc(ByVal ItemNum As Long)
    Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT, destRECT As D3DRECT
    Dim ItemPic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
        ItemPic = Item(ItemNum).Pic

        If ItemPic = 0 Then Exit Sub
        
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene

        With rec
            .Top = 0
            .Bottom = .Top + PIC_Y
            .Left = Tex_Item(ItemPic).Width / 2
            .Right = .Left + PIC_X
        End With

        With rec_pos
            .Top = 0
            .Bottom = 64
            .Left = 0
            .Right = 64
        End With
        
        RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos

        With destRECT
            .X1 = 0
            .Y1 = 0
            .Y2 = 64
            .X2 = 64
        End With
        
        Direct3D_Device.EndScene
        Direct3D_Device.Present destRECT, destRECT, frmMain.picItemDescPic.hwnd, ByVal (0)
    End If
    Exit Sub

' Error handler
errorhandler:
    HandleError "DrawItemDesc", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawSpellDesc(ByVal SpellNum As Long)
    Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT, destRECT As D3DRECT
    Dim SpellPic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If SpellNum > 0 And SpellNum <= MAX_SPELLS Then
        SpellPic = Spell(SpellNum).Icon

        If SpellPic <= 0 Or SpellPic > NumSpellIcons Then Exit Sub
        
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene

        With rec
            .Top = 0
            .Bottom = .Top + PIC_Y
            .Left = 0
            .Right = .Left + PIC_X
        End With

        With rec_pos
            .Top = 0
            .Bottom = 64
            .Left = 0
            .Right = 64
        End With
        
        RenderTextureByRects Tex_SpellIcon(SpellPic), rec, rec_pos

        With destRECT
            .X1 = 0
            .Y1 = 0
            .Y2 = 64
            .X2 = 64
        End With
        
        Direct3D_Device.EndScene
        Direct3D_Device.Present destRECT, destRECT, frmMain.picSpellDescPic.hwnd, ByVal (0)
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawSpellDesc", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawTileOutline()
    Dim rec As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Map.OptBlock.Value Then Exit Sub

    With rec
        .Top = 0
        .Bottom = .Top + PIC_Y
        .Left = 0
        .Right = .Left + PIC_X
    End With

    RenderTexture Tex_Misc, ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawTileOutline", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub NewCharacterDrawSprite()
    Dim Sprite As Long, srcRect As D3DRECT, destRECT As D3DRECT
    Dim sRect As RECT
    Dim dRect As RECT
    Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If frmMenu.cmbClass.ListIndex = -1 Then Exit Sub
    
    If frmMenu.optMale.Value = True Then
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).MaleSprite
    Else
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).FemaleSprite
    End If
    
    If Sprite < 1 Or Sprite > NumCharacters Then
        frmMenu.picSprite.Cls
        Exit Sub
    End If
    
    Width = Tex_Character(Sprite).Width / 4
    Height = Tex_Character(Sprite).Height / 4
    
    frmMenu.picSprite.Width = Width
    frmMenu.picSprite.Height = Height
    
    sRect.Top = 0
    sRect.Bottom = sRect.Top + Height
    sRect.Left = 0
    sRect.Right = sRect.Left + Width
    
    dRect.Top = 0
    dRect.Bottom = Height
    dRect.Left = 0
    dRect.Right = Width
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    RenderTextureByRects Tex_Character(Sprite), sRect, dRect
    
    With srcRect
        .X1 = 0
        .X2 = Width
        .Y1 = 0
        .Y2 = Height
    End With
                    
    With destRECT
        .X1 = 0
        .X2 = Width
        .Y1 = 0
        .Y2 = Height
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, destRECT, frmMenu.picSprite.hwnd, ByVal (0)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "NewCharacterDrawSprite", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub Render_Graphics()
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim rec As RECT
    Dim rec_pos As RECT, srcRect As D3DRECT
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    ' Check for device lost
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then HandleDeviceLost: Exit Sub
    
    ' Don't render
    If frmMain.WindowState = vbMinimized Then Exit Sub
    
    If GettingMap Then Exit Sub
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorARGB(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    ' Update the viewpoint
    Call UpdateCamera
    
    ' Update draw Name
    UpdateDrawMapName
    
    If Map.Panorama > 0 And Map.Panorama <= NumPanoramas Then
        RenderTexture Tex_Panorama(Map.Panorama), 0, 0, 0, 0, Tex_Panorama(Map.Panorama).Width, Tex_Panorama(Map.Panorama).Height, Tex_Panorama(Map.Panorama).Width, Tex_Panorama(Map.Panorama).Height, -1
    End If
    
    ' Draw lower tiles
    If NumTileSets > 0 Then
        For x = TileView.Left To TileView.Right
            For y = TileView.Top To TileView.Bottom
                If IsValidMapPoint(x, y) Then
                    Call DrawMapLowerTiles(x, y)
                End If
            Next
        Next
    End If
    
    ' Render the decals
    If Options.Blood = 1 Then
        For i = 1 To Blood_HighIndex
            Call DrawBlood(i)
        Next
    End If

    ' Draw out the items
    If NumItems > 0 Then
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).Num > 0 Then
                Call DrawMapItem(i)
            End If
        Next
    End If
    
    ' Draw out lower events
    If Map.CurrentEvents > 0 Then
        For i = 1 To Map.CurrentEvents
            If Map.MapEvents(i).Position = 0 Then
                DrawEvent i
            End If
        Next
    End If
    
    ' Draw animations
    If NumAnimations > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(0) Then
                DrawAnimation i, 0
            End If
        Next
    End If

    ' Y-based render. Renders players, npcs, and resources based on Y-axis.
    For y = TileView.Top To TileView.Bottom
        ' Npcs
        For i = 1 To Map.Npc_HighIndex
            If MapNPC(i).y = y Then
                Call DrawNpc(i)
            End If
        Next
        
        ' Players
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                If Player(i).y = y Then
                    If Not i = MyIndex Then
                        Call DrawPlayer(i)
                    End If
                End If
            End If
        Next
        
        ' Render our sprite now so it's always at the top
        If Player(MyIndex).y = y Then
            Call DrawPlayer(MyIndex)
        End If
        
        ' Events
        If Map.CurrentEvents > 0 Then
            For i = 1 To Map.CurrentEvents
                If Map.MapEvents(i).Position = 1 Then
                    If y = Map.MapEvents(i).y Then
                        DrawEvent i
                    End If
                End If
            Next
        End If
        
        ' Resources
        If NumResources > 0 Then
            If Resources_Init Then
                If Resource_Index > 0 Then
                    For i = 1 To Resource_Index
                        If MapResource(i).y = y Then
                            Call DrawMapResource(i)
                        End If
                    Next
                End If
            End If
        End If
    Next
    
    ' Animations
    If NumAnimations > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(1) Then
                DrawAnimation i, 1
            End If
        Next
    End If

    ' Draw out upper tiles
    If NumTileSets > 0 Then
        For x = TileView.Left To TileView.Right
            For y = TileView.Top To TileView.Bottom
                If IsValidMapPoint(x, y) Then
                    Call DrawMapUpperTiles(x, y)
                End If
            Next
        Next
    End If
    
    ' Draw out higher events
    If Map.CurrentEvents > 0 Then
        For i = 1 To Map.CurrentEvents
            If Map.MapEvents(i).Position = 2 Then
                DrawEvent i
            End If
        Next
    End If
    
    DrawWeather
    DrawFog
    DrawTint
    
    ' Draw the bars
    Call DrawBars
    
    ' Draw out a square at the mouse cursor
    If InMapEditor Then
        If frmEditor_Map.OptBlock.Value Then
            For x = TileView.Left To TileView.Right
                For y = TileView.Top To TileView.Bottom
                    If IsValidMapPoint(x, y) Then
                        Call DrawGrid(x, y)
                        Call DrawDirection(x, y)
                    End If
                Next
            Next
        ElseIf frmEditor_Map.chkGrid Then
            For x = TileView.Left To TileView.Right
                For y = TileView.Top To TileView.Bottom
                    If IsValidMapPoint(x, y) Then
                        Call DrawGrid(x, y)
                    End If
                Next
            Next
        End If
    End If
    
    ' Draw the target icon
    If MyTarget > 0 Then
        If MyTargetType = TARGET_TYPE_PLAYER Then
            DrawTarget (Player(MyTarget).x * 32) + TempPlayer(MyTarget).xOffset, (Player(MyTarget).y * 32) + TempPlayer(MyTarget).yOffset
        ElseIf MyTargetType = TARGET_TYPE_NPC Then
            DrawTarget (MapNPC(MyTarget).x * 32) + MapNPC(MyTarget).xOffset, (MapNPC(MyTarget).y * 32) + MapNPC(MyTarget).yOffset
        End If
    End If
    
    ' Draw the hover icon
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Player(i).Map = Player(MyIndex).Map Then
                If CurX = Player(i).x And CurY = Player(i).y Then
                    If MyTargetType = TARGET_TYPE_PLAYER And MyTarget = i Then
                        ' Don't render
                    Else
                        DrawHover TARGET_TYPE_PLAYER, i, (Player(i).x * 32) + TempPlayer(i).xOffset, (Player(i).y * 32) + TempPlayer(i).yOffset
                    End If
                End If
            End If
        End If
    Next
    
    For i = 1 To Map.Npc_HighIndex
        If MapNPC(i).Num > 0 Then
            If CurX = MapNPC(i).x And CurY = MapNPC(i).y Then
                If MyTargetType = TARGET_TYPE_NPC And MyTarget = i Then
                    ' Don't render
                Else
                    DrawHover TARGET_TYPE_NPC, i, (MapNPC(i).x * 32) + MapNPC(i).xOffset, (MapNPC(i).y * 32) + MapNPC(i).yOffset
                End If
            End If
        End If
    Next
    
    ' Draw weater
    If DrawThunder > 0 Then RenderTexture Tex_White, 0, 0, 0, 0, frmMain.picScreen.ScaleWidth, frmMain.picScreen.ScaleHeight, 32, 32, D3DColorRGBA(255, 255, 255, 160): DrawThunder = DrawThunder - 1
    
    ' Get rec
    With rec
        .Top = Camera.Top
        .Bottom = .Top + ScreenY
        .Left = Camera.Left
        .Right = .Left + ScreenX
    End With
        
    ' rec_pos
    With rec_pos
        .Bottom = ScreenY
        .Right = ScreenX
    End With
        
    With srcRect
        .X1 = 0
        .X2 = frmMain.picScreen.ScaleWidth
        .Y1 = 0
        .Y2 = frmMain.picScreen.ScaleHeight
    End With
    
    If InMapEditor Then Call DrawMapAttributes
    
    ' Draw player names
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
            Call DrawPlayerName(i)
        End If
    Next
    
    For i = 1 To Map.CurrentEvents
        If Map.MapEvents(i).Visible = 1 Then
            If Map.MapEvents(i).ShowName = 1 Then
                Call DrawEventName(i)
            End If
        End If
    Next
    
    ' Draw npc names
    For i = 1 To Map.Npc_HighIndex
        If MapNPC(i).Num > 0 Then
            Call DrawNPCName(i)
        End If
    Next
    
    ' draw the messages
    For i = 1 To ChatBubble_HighIndex
        If ChatBubble(i).active Then
            Call DrawChatBubble(i)
        End If
    Next
    
    ' Draw emotions
    DrawEmoticons
    
    ' Draw action messages
    For i = 1 To Action_HighIndex
        Call DrawActionMsg(i)
    Next

    ' Draw map name
    RenderText Font_Default, Map.name, DrawMapNameX, DrawMapNameY, DrawMapNameColor
    
    If InMapEditor And frmEditor_Map.OptEvents.Value Then DrawEvents
    If InMapEditor And frmEditor_Map.OptLayers Then DrawTileOutline

    If FadeAmount > 0 Then RenderTexture Tex_Fade, 0, 0, 0, 0, frmMain.picScreen.ScaleWidth, frmMain.picScreen.ScaleHeight, 32, 32, D3DColorRGBA(255, 255, 255, FadeAmount)
    If FlashTimer > timeGetTime Then RenderTexture Tex_White, 0, 0, 0, 0, frmMain.picScreen.ScaleWidth, frmMain.picScreen.ScaleHeight, 32, 32, -1
    
    ' Draw loc
    If BLoc Then
        RenderText Font_Default, Trim$("Cur X: " & CurX & " Y: " & CurY), 8, 85, Yellow
        RenderText Font_Default, Trim$("Loc X: " & GetPlayerX(MyIndex) & " y: " & GetPlayerY(MyIndex)), 8, 100, Yellow
        RenderText Font_Default, Trim$(" (Map #" & GetPlayerMap(MyIndex) & ")"), 8, 115, Yellow
    End If
    
    ' Draw fps
    If BFPS Then
        If FPS_Lock Then
            If GUIVisible Then
                RenderText Font_Default, "FPS: " & Round(GameFPS / 1500) & " Ping: " & CStr(Ping), 300, 48, White
            Else
                RenderText Font_Default, "FPS: " & Round(GameFPS / 1500) & " Ping: " & CStr(Ping), 300, 8, White
            End If
        Else
            If GUIVisible Then
                RenderText Font_Default, "FPS: " & GameFPS & " Ping: " & CStr(Ping), 300, 48, White
            Else
                RenderText Font_Default, "FPS: " & GameFPS & " Ping: " & CStr(Ping), 300, 8, White
            End If
        End If
    End If
    
    Call Direct3D_Device.EndScene
    Call Direct3D_Device.Present(ByVal 0, ByVal 0, 0, ByVal 0)
        
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost
        Exit Sub
    Else
        If InShop = 0 And InBank = False Then Direct3D_Device.Present srcRect, ByVal 0, 0, ByVal 0
        DrawGDI
    End If
    Exit Sub

' Error handler
errorhandler:
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost
        Exit Sub
    Else
        If Options.Debug = 1 Then
            HandleError "Render_Graphics", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
            Err.Clear
        End If
        MsgBox "Unrecoverable DX8 error."
        DestroyGame
    End If
End Sub

Public Function ConvertMapX(ByVal x As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ConvertMapX = x - (TileView.Left * PIC_X) - Camera.Left
    Exit Function
    
' Error handler
errorhandler:
    HandleError "ConvertMapX", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Function ConvertMapY(ByVal y As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ConvertMapY = y - (TileView.Top * PIC_Y) - Camera.Top
    Exit Function
    
' Error handler
errorhandler:
    HandleError "ConvertMapY", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Function InViewPort(ByVal x As Long, ByVal y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If x < TileView.Left Then Exit Function
    If y < TileView.Top Then Exit Function
    If x > TileView.Right Then Exit Function
    If y > TileView.Bottom Then Exit Function
    
    InViewPort = True
    Exit Function
    
' Error handler
errorhandler:
    HandleError "InViewPort", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Function IsValidMapPoint(ByVal x As Long, ByVal y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If x < 0 Then Exit Function
    If y < 0 Then Exit Function
    If x > Map.MaxX Then Exit Function
    If y > Map.MaxY Then Exit Function
    
    IsValidMapPoint = True
    Exit Function
    
' Error handler
errorhandler:
    HandleError "IsValidMapPoint", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Sub LoadTilesets()
    Dim x As Long
    Dim y As Long
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim TilesetInUse(0 To NumTileSets)
    
    For x = 0 To Map.MaxX
        For y = 0 To Map.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                ' Check exists
                If Map.Tile(x, y).Layer(i).Tileset > 0 And Map.Tile(x, y).Layer(i).Tileset <= NumTileSets Then
                    TilesetInUse(Map.Tile(x, y).Layer(i).Tileset) = True
                End If
            Next
        Next
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "LoadTilesets", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub DrawBank()
    Dim i As Long, x As Long, y As Long, ItemNum As Long, srcRect As D3DRECT, destRECT As D3DRECT
    Dim Amount As String
    Dim sRect As RECT, dRect As RECT
    Dim Sprite As Long, Color As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmMain.picBank.Visible Then
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
    
        For i = 1 To MAX_BANK
            ItemNum = GetBankItemNum(i)
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            
                Sprite = Item(ItemNum).Pic
                
                If Sprite <= 0 Or Sprite > NumItems Then Exit Sub
                    If Tex_Item(Sprite).Width <= 64 Then
                        With sRect
                            .Top = 0
                            .Bottom = .Top + PIC_Y
                            .Left = Tex_Item(Sprite).Width / 2
                            .Right = .Left + PIC_X
                        End With
                        
                        With dRect
                            .Top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                            .Bottom = .Top + PIC_Y
                            .Left = BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                            .Right = .Left + PIC_X
                        End With
                    
                    RenderTextureByRects Tex_Item(Sprite), sRect, dRect
    
                    ' If item is a stack - draw the amount you have
                    If GetBankItemValue(i) > 1 Then
                        y = dRect.Top + 22
                        x = dRect.Left - 4
                        Amount = GetBankItemValue(i)
                        
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If CLng(Amount) < 1000000 Then
                            Color = White
                        ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                            Color = Yellow
                        ElseIf CLng(Amount) > 10000000 Then
                            Color = BrightGreen
                        End If
                        RenderText Font_Default, ConvertCurrency(Amount), x, y, Color
                    End If
                End If
            End If
        Next
    
        With srcRect
            .X1 = BankLeft
            .X2 = .X1 + 400
            .Y1 = BankTop
            .Y2 = .Y1 + 310
        End With
                    
        With destRECT
            .X1 = BankLeft
            .X2 = .X1 + 400
            .Y1 = BankTop
            .Y2 = 310 + .Y1
        End With
                    
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRECT, frmMain.picBank.hwnd, ByVal (0)
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawBank", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawBankItem(ByVal x As Long, ByVal y As Long)
    Dim sRect As RECT, dRect As RECT, srcRect As D3DRECT, destRECT As D3DRECT
    Dim ItemNum As Long
    Dim Sprite As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ItemNum = GetBankItemNum(DragBankSlot)
    Sprite = Item(GetBankItemNum(DragBankSlot)).Pic
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 255), 1#, 0
    Direct3D_Device.BeginScene
    
    If ItemNum > 0 Then
        If ItemNum <= MAX_ITEMS Then
            With sRect
                .Top = 0
                .Bottom = .Top + PIC_Y
                .Left = Tex_Item(Sprite).Width / 2
                .Right = .Left + PIC_X
            End With
        End If
    End If
    
    With dRect
        .Top = 2
        .Bottom = .Top + PIC_Y
        .Left = 2
        .Right = .Left + PIC_X
    End With

    RenderTextureByRects Tex_Item(Sprite), sRect, dRect
    
    With frmMain.picTempBank
        .Top = y
        .Left = x
        .Visible = True
        .ZOrder (0)
    End With
    
    With srcRect
        .X1 = 0
        .X2 = 32
        .Y1 = 0
        .Y2 = 32
    End With
    
    With destRECT
        .X1 = 2
        .Y1 = 2
        .Y2 = .Y1 + 32
        .X2 = .X1 + 32
    End With
    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, destRECT, frmMain.picTempBank.hwnd, ByVal (0)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawBankItem", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawAutoTile(ByVal layerNum As Long, ByVal destX As Long, ByVal destY As Long, ByVal quarterNum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal Alpha As Byte = 255)
    Dim yOffset As Long, xOffset As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Calculate the offset
    Select Case Map.Tile(x, y).Autotile(layerNum)
        Case AUTOTILE_WATERFALL
            yOffset = (waterfallFrame - 1) * 32
        Case AUTOTILE_ANIM
            xOffset = autoTileFrame * 64
        Case AUTOTILE_CLIFF
            yOffset = -32
    End Select
    
    ' Draw the quarter
    RenderTexture Tex_Tileset(Map.Tile(x, y).Layer(layerNum).Tileset), destX, destY, Autotile(x, y).Layer(layerNum).srcX(quarterNum) + xOffset, Autotile(x, y).Layer(layerNum).srcY(quarterNum) + yOffset, 16, 16, 16, 16, D3DColorARGB(Alpha, 255, 255, 255)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawAutoTile", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorMap_DrawRandom()
    Dim sRect As RECT
    Dim dRect As RECT
    Dim x As Long, y As Long
    Dim i As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    For i = 0 To 3
        If RandomTileSheet(i) = 0 Then
            Exit Sub
        End If
        
        x = RandomTile(i) Mod 16
        y = (RandomTile(i) - x) / 16
        
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
        
        sRect.Top = y * PIC_Y
        sRect.Bottom = sRect.Top + PIC_Y
        sRect.Left = x * PIC_X
        sRect.Right = sRect.Left + PIC_X
        
        dRect = sRect
        dRect.Top = 0
        dRect.Bottom = PIC_Y
        dRect.Left = 0
        dRect.Right = PIC_X
        
        RenderTextureByRects Tex_Tileset(RandomTileSheet(i)), sRect, dRect
    
        Direct3D_Device.EndScene
        Direct3D_Device.Present dRect, dRect, frmEditor_Map.picRandomTile(i).hwnd, ByVal (0)
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "EditorMap_DrawRandom", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' Character Editor
Public Sub EditorChar_AnimSprite()
    Dim srcRect As D3DRECT, destRECT As D3DRECT
    Dim sRect As RECT
    Dim dRect As RECT
    Dim x As Byte, y As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If CharSpritePos > 15 Then CharSpritePos = 0
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    x = CharSpritePos Mod 4
    y = (CharSpritePos - x) / 4
    
    sRect.Top = y * 48
    sRect.Bottom = sRect.Top + 48
    sRect.Left = x * 32
    sRect.Right = sRect.Left + 32

    dRect = sRect
    dRect.Top = 0
    dRect.Bottom = 48
    dRect.Left = 0
    dRect.Right = 32
    
    RenderTextureByRects Tex_CharSprite, sRect, dRect
          
    With destRECT
        .X1 = 0
        .X2 = frmCharEditor.picSprite.ScaleWidth
        .Y1 = 0
        .Y2 = frmCharEditor.picSprite.ScaleHeight
    End With

    Direct3D_Device.EndScene
    Direct3D_Device.Present destRECT, destRECT, frmCharEditor.picSprite.hwnd, ByVal (0)

    CharSpritePos = CharSpritePos + 1
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "EditorChar_AnimSprite", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorClass_DrawFace(ByVal Gender As Byte)
    Dim Sprite As Long
    Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If NumFaces = 0 Then Exit Sub
    
    If Gender = 0 Then
        Sprite = frmEditor_Class.scrlMFace.Value
    Else
        Sprite = frmEditor_Class.scrlFFace.Value
    End If
    
    If Sprite <= 0 Or Sprite > NumFaces Then
        frmEditor_Class.picFace.Cls
        Exit Sub
    End If
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene

    With rec
        .Bottom = Tex_Face(Sprite).Height
        .Left = 0
        .Right = Tex_Face(Sprite).Width
    End With

    With rec_pos
        .Top = 0
        .Bottom = Tex_Face(Sprite).Height
        .Left = 0
        .Right = Tex_Face(Sprite).Width
    End With

    RenderTextureByRects Tex_Face(Sprite), rec, rec_pos
    
    With srcRect
        .X1 = 0
        .X2 = frmEditor_Class.picFace.Width
        .Y1 = 0
        .Y2 = frmEditor_Class.picFace.Height
    End With
    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, srcRect, frmEditor_Class.picFace.hwnd, ByVal (0)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "EditorClass_DrawFace", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub DrawEventChatFace()
    Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If NumFaces = 0 Then Exit Sub
    
    If EventFace <= 0 Or EventFace > NumFaces Then Exit Sub
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene

    With rec
        .Top = 0
        .Bottom = Tex_Face(EventFace).Height
        .Left = 0
        .Right = Tex_Face(EventFace).Width
    End With

    With rec_pos
        .Top = 0
        .Bottom = Tex_Face(EventFace).Height
        .Left = 0
        .Right = Tex_Face(EventFace).Width
    End With

    RenderTextureByRects Tex_Face(EventFace), rec, rec_pos
    
    With srcRect
        .X1 = 0
        .X2 = frmMain.picChatFace.Width
        .Y1 = 0
        .Y2 = frmMain.picChatFace.Height
    End With
    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, srcRect, frmMain.picChatFace.hwnd, ByVal (0)
    
    frmMain.picChatFace.Height = Tex_Face(EventFace).Height
    frmMain.picChatFace.Width = Tex_Face(EventFace).Width
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawEventChatFace", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub EditorEvent_DrawFace()
    Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT
    Dim FaceNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If NumFaces = 0 Then Exit Sub
    
    FaceNum = frmEditor_Events.scrlFace.Value
    
    If FaceNum <= 0 Or FaceNum > NumFaces Then
        frmEditor_Events.picFace.Cls
        Exit Sub
    End If
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene

    With rec
        .Top = 0
        .Bottom = Tex_Face(FaceNum).Height
        .Left = 0
        .Right = Tex_Face(FaceNum).Width
    End With

    With rec_pos
        .Top = 0
        .Bottom = Tex_Face(FaceNum).Height
        .Left = 0
        .Right = Tex_Face(FaceNum).Width
    End With

    RenderTextureByRects Tex_Face(FaceNum), rec, rec_pos
    
    With srcRect
        .X1 = 0
        .X2 = frmEditor_Events.picFace.Width
        .Y1 = 0
        .Y2 = frmEditor_Events.picFace.Height
    End With
    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, srcRect, frmEditor_Events.picFace.hwnd, ByVal (0)
    
    frmEditor_Events.picFace.Height = PixelsToTwips(Tex_Face(FaceNum).Height, 1)
    frmEditor_Events.picFace.Width = PixelsToTwips(Tex_Face(FaceNum).Width, 0)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "EditorEvent_DrawFace", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub EditorEvent_DrawFace2()
    Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT
    Dim FaceNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If NumFaces = 0 Then Exit Sub
    
    FaceNum = frmEditor_Events.scrlFace2.Value
    
    If FaceNum <= 0 Or FaceNum > NumFaces Then
        frmEditor_Events.picFace2.Cls
        Exit Sub
    End If
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene

    With rec
        .Top = 0
        .Bottom = Tex_Face(FaceNum).Height
        .Left = 0
        .Right = Tex_Face(FaceNum).Width
    End With

    With rec_pos
        .Top = 0
        .Bottom = Tex_Face(FaceNum).Height
        .Left = 0
        .Right = Tex_Face(FaceNum).Width
    End With

    RenderTextureByRects Tex_Face(FaceNum), rec, rec_pos
    
    With srcRect
        .X1 = 0
        .X2 = frmEditor_Events.picFace2.Width
        .Y1 = 0
        .Y2 = frmEditor_Events.picFace2.Height
    End With
    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, srcRect, frmEditor_Events.picFace2.hwnd, ByVal (0)
    
    frmEditor_Events.picFace2.Height = PixelsToTwips(Tex_Face(FaceNum).Height, 1)
    frmEditor_Events.picFace2.Width = PixelsToTwips(Tex_Face(FaceNum).Width, 0)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "EditorEvent_DrawFace2", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Function IsConstAnimated(ByVal Sprite As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If AnimatedSpriteNumbers = vbNullString Then Exit Function
    
    If AnimatedSprites(Sprite) = 1 Then
        IsConstAnimated = True
        Exit Function
    End If
    
    IsConstAnimated = False
    Exit Function
    
' Error handler
errorhandler:
    HandleError "IsConstAnimated", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Sub EditorMap_DrawTilePreview()
    Dim Height As Long
    Dim Width As Long
    Dim x As Long
    Dim y As Long
    Dim Tileset As Long
    Dim srcRect As RECT
    Dim destRECT As D3DRECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not IsInBounds Then Exit Sub
    
    ' Find tileset number
    Tileset = frmEditor_Map.scrlTileSet.Value
    
    x = CurX * PIC_X
    y = CurY * PIC_Y
    
    Height = EditorTileHeight
    Width = EditorTileWidth
    
    With srcRect
        .Left = 0
        .Top = 0
        .Right = srcRect.Left - Width
        .Bottom = srcRect.Top - Height
    End With
    
    With destRECT
        .X1 = (EditorTileX * PIC_X) - srcRect.Left
        .X2 = (EditorTileWidth * PIC_X) + .X1
        .Y1 = (EditorTileY * PIC_Y) - srcRect.Top
        .Y2 = (EditorTileHeight * PIC_Y) + .Y1
    End With
    
    RenderTexture Tex_Tileset(Tileset), ConvertMapX(x), ConvertMapY(y), destRECT.X1, destRECT.Y1, Width * PIC_X, Height * PIC_Y, Width * PIC_X, Height * PIC_Y, D3DColorARGB(4, 255, 255, 255)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawMapTilesPreview", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorClass_DrawSprite(ByVal Gender As Byte)
    Dim Sprite As Integer
    Dim sRect As RECT
    Dim dRect As RECT
    Dim destRECT As D3DRECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Gender = 0 Then
        Sprite = frmEditor_Class.scrlMSprite.Value
    Else
        Sprite = frmEditor_Class.scrlFSprite.Value
    End If

    If Sprite < 1 Or Sprite > NumCharacters Then
        frmEditor_Class.picSprite.Cls
        Exit Sub
    End If
    
    sRect.Top = 0
    sRect.Bottom = Tex_Character(Sprite).Height / 4
    sRect.Left = (Tex_Character(Sprite).Width / 4) * 2 ' Facing down
    sRect.Right = sRect.Left + Tex_Character(Sprite).Width / 4
    dRect.Top = 0
    dRect.Bottom = Tex_Character(Sprite).Height / 4
    dRect.Left = 0
    dRect.Right = Tex_Character(Sprite).Width / 4
    
    frmEditor_Class.picSprite.Width = PixelsToTwips(Tex_Character(Sprite).Width / 4, 0)
    frmEditor_Class.picSprite.Height = PixelsToTwips(Tex_Character(Sprite).Height / 4, 1)

    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Character(Sprite), sRect, dRect
    
    With destRECT
        .X1 = 0
        .X2 = frmEditor_Class.picSprite.Width
        .Y1 = 0
        .Y2 = frmEditor_Class.picSprite.Height
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRECT, destRECT, frmEditor_Class.picSprite.hwnd, ByVal (0)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "EditorClass_DrawSprite", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub RenderOptionButton(ByRef ThePictureBox As PictureBox, ByVal TheOption As Byte, ByVal TheValue As Byte)
    Dim FileName As String

    If TheValue = 0 Then
        FileName = App.Path & GFX_PATH & "gui/main/buttons/option_off.bmp"
    ElseIf TheValue = 1 Then
        FileName = App.Path & GFX_PATH & "gui/main/buttons/option_on.bmp"
    ElseIf TheValue = 2 Then
        FileName = App.Path & GFX_PATH & "gui/main/buttons/option_off_hover.bmp"
    ElseIf TheValue = 3 Then
        FileName = App.Path & GFX_PATH & "gui/main/buttons/option_on_hover.bmp"
    End If
    
    OptionButton(TheOption).State = TheValue
    ThePictureBox.Picture = LoadPicture(FileName)
End Sub

Public Sub ResizeHPBar()
    Dim MoveSpeed As Long
        
    If Not NewHPBarWidth = OldHPBarWidth Then
        If NewHPBarWidth > OldHPBarWidth Then
            MoveSpeed = NewHPBarWidth - OldHPBarWidth
            If MoveSpeed < 25 Then MoveSpeed = 25
            
            If OldHPBarWidth - (MoveSpeed / 25) > NewHPBarWidth Then
                frmMain.imgHPBar.Width = NewHPBarWidth
            Else
                frmMain.imgHPBar.Width = OldHPBarWidth + (MoveSpeed / 25)
            End If
        Else
            MoveSpeed = OldHPBarWidth - NewHPBarWidth
            If MoveSpeed < 25 Then MoveSpeed = 25
            
            If NewHPBarWidth + (MoveSpeed / 25) > OldHPBarWidth Then
                frmMain.imgHPBar.Width = NewHPBarWidth
            Else
                frmMain.imgHPBar.Width = OldHPBarWidth - (MoveSpeed / 25)
            End If
        End If
        
        OldHPBarWidth = frmMain.imgHPBar.Width
    End If
End Sub

Public Sub ResizeMPBar()
    Dim MoveSpeed As Long
    
    If Not NewMPBarWidth = OldMPBarWidth Then
        If NewMPBarWidth > OldMPBarWidth Then
            MoveSpeed = NewMPBarWidth - OldMPBarWidth
            If MoveSpeed < 25 Then MoveSpeed = 25
            
            If OldMPBarWidth - (MoveSpeed / 25) > NewMPBarWidth Then
                frmMain.imgMPBar.Width = NewMPBarWidth
            Else
                frmMain.imgMPBar.Width = OldMPBarWidth + (MoveSpeed / 25)
            End If
        Else
            MoveSpeed = OldMPBarWidth - NewMPBarWidth
            If MoveSpeed < 25 Then MoveSpeed = 25
            
            If NewMPBarWidth + (MoveSpeed / 25) > OldMPBarWidth Then
                frmMain.imgMPBar.Width = NewMPBarWidth
            Else
                frmMain.imgMPBar.Width = OldMPBarWidth - (MoveSpeed / 25)
            End If
        End If
        
        OldMPBarWidth = frmMain.imgMPBar.Width
    End If
End Sub

Public Sub ResizeExpBar()
    Dim MoveSpeed As Long
    
    If Not NewEXPBarWidth = OldEXPBarWidth Then
        If NewEXPBarWidth > OldEXPBarWidth Then
            MoveSpeed = NewEXPBarWidth - OldEXPBarWidth
            If MoveSpeed < 25 Then MoveSpeed = 25
            
            If OldEXPBarWidth - (MoveSpeed / 25) > NewEXPBarWidth Then
                frmMain.imgEXPBar.Width = NewEXPBarWidth
            Else
                frmMain.imgEXPBar.Width = OldEXPBarWidth + (MoveSpeed / 25)
            End If
        Else
            MoveSpeed = OldEXPBarWidth - NewEXPBarWidth
            If MoveSpeed < 25 Then MoveSpeed = 25
            
            If NewEXPBarWidth + (MoveSpeed / 25) > OldEXPBarWidth Then
                frmMain.imgEXPBar.Width = NewEXPBarWidth
            Else
                frmMain.imgEXPBar.Width = OldEXPBarWidth - (MoveSpeed / 25)
            End If
        End If
        
        OldEXPBarWidth = frmMain.imgEXPBar.Width
    End If
End Sub

Public Sub DrawEquipment()
    Dim i As Long
    Dim ItemNum As Long
    Dim ItemPic As Long
    Dim sRect As RECT
    Dim dRect As RECT

    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    With sRect
        .Top = 0
        .Bottom = Tex_Equip.Height
        .Left = 0
        .Right = Tex_Equip.Width
    End With
    
    With dRect
        .Top = 0
        .Bottom = frmMain.picEquipment.Height
        .Left = 0
        .Right = frmMain.picEquipment.Width
    End With

    RenderTextureByRects Tex_Equip, sRect, dRect
    
    ' Now lets make the image that we will be rendering today
    For i = 1 To Equipment.Equipment_Count - 1
        ItemNum = GetPlayerEquipment(MyIndex, i)
        
        ' If there is an item draw it, if not do NOTHING!
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ItemPic = Item(ItemNum).Pic
            
            ' If the picture exists then render it
            If ItemPic > 0 And ItemPic <= NumItems Then
                sRect.Top = 0
                sRect.Bottom = PIC_Y
                sRect.Left = 0
                sRect.Right = PIC_X

                RenderTexture Tex_Item(ItemPic), EquipSlotLeft(i), EquipSlotTop(i), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top
            End If
        End If
    Next
    
    With sRect
        .Top = 0
        .Bottom = Tex_Equip.Height
        .Left = 0
        .Right = Tex_Equip.Width
    End With
    
    With dRect
        .Top = 0
        .Bottom = frmMain.picEquipment.Height
        .Left = 0
        .Right = frmMain.picEquipment.Width
    End With
    
    Direct3D_Device.EndScene
    Direct3D_Device.Present sRect, dRect, frmMain.picEquipment.hwnd, ByVal (0)
End Sub

Public Sub EditorEmoticon_DrawIcon()
    Dim EmoticonNum As Long
    Dim sRect As RECT
    Dim dRect As RECT
    Dim destRECT As D3DRECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    EmoticonNum = frmEditor_Emoticon.scrlEmoticon.Value

    If EmoticonNum < 1 Or EmoticonNum > NumEmoticons Then
        frmEditor_Emoticon.picEmoticon.Cls
        Exit Sub
    End If

    sRect.Top = 0
    sRect.Bottom = Tex_Emoticon(EmoticonNum).Height
    sRect.Left = 0
    sRect.Right = sRect.Left + Tex_Emoticon(EmoticonNum).Width
    dRect.Top = 0
    dRect.Bottom = Tex_Emoticon(EmoticonNum).Height
    dRect.Left = 0
    dRect.Right = Tex_Emoticon(EmoticonNum).Width
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Emoticon(EmoticonNum), sRect, dRect
    
    With destRECT
        .X1 = 0
        .X2 = frmEditor_Emoticon.picEmoticon.Width
        .Y1 = 0
        .Y2 = frmEditor_Emoticon.picEmoticon.Height
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRECT, destRECT, frmEditor_Emoticon.picEmoticon.hwnd, ByVal (0)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "EditorEmoticon_BltIcon", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawEmoticons()
    Dim sRect As RECT
    Dim dRect As RECT
    Dim EmoticonNum As Byte, i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(MyIndex) = GetPlayerMap(i) Then
            EmoticonNum = TempPlayer(i).EmoticonNum
            
            If EmoticonNum < 1 Or EmoticonNum > NumEmoticons Then
                If Trim$(Player(MyIndex).Status) = "AFK" Then
                    EmoticonNum = Emoticon(1).Pic
                Else
                    Exit Sub
                End If
            End If
            
            ' Clear out the data if it needs to disappear
            If timeGetTime > TempPlayer(i).EmoticonTimer And EmoticonNum <> Emoticon(1).Pic Then
                TempPlayer(i).EmoticonNum = 0
                TempPlayer(i).EmoticonTimer = 0
                Exit Sub
            End If
    
            If InViewPort(GetPlayerX(i), GetPlayerY(i)) Then
                With sRect
                    .Top = 0
                    .Bottom = .Top + PIC_Y
                    .Left = 0
                    .Right = .Left + PIC_X
                End With
                
                ' Same for destination as source
                dRect = sRect
                
                RenderTexture Tex_Emoticon(EmoticonNum), GetPlayerTextX(i) - 16, GetPlayerTextY(i) - 16, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
            End If
        End If
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawEmoticons", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorAnim_DrawAnim()
    Dim Animationnum As Long
    Dim sRect As RECT
    Dim dRect As RECT
    Dim i As Long
    Dim Width As Long, Height As Long, srcRect As D3DRECT, destRECT As D3DRECT
    Dim LoopTime As Long
    Dim FrameCount As Long
    Dim ShouldRender As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 0 To 1
        Animationnum = frmEditor_Animation.scrlSprite(i).Value
        
        If Animationnum < 1 Or Animationnum > NumAnimations Then
            frmEditor_Animation.picSprite(i).Cls
        Else
            LoopTime = frmEditor_Animation.scrlLoopTime(i)
            FrameCount = frmEditor_Animation.scrlFrameCount(i)
            
            ShouldRender = False
            
            ' Check if we need to render new frame
            If AnimEditorTimer(i) + LoopTime <= timeGetTime Then
                ' Check if out of range
                If AnimEditorFrame(i) >= FrameCount Then
                    AnimEditorFrame(i) = 1
                Else
                    AnimEditorFrame(i) = AnimEditorFrame(i) + 1
                End If
                
                AnimEditorTimer(i) = timeGetTime
                ShouldRender = True
            End If
        
            If ShouldRender Then
                If frmEditor_Animation.scrlFrameCount(i).Value > 0 Then
                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    
                    ' Total width divided by frame count
                    Width = Tex_Animation(Animationnum).Width / frmEditor_Animation.scrlFrameCount(i).Value
                    Height = Tex_Animation(Animationnum).Height
                    
                    sRect.Top = 0
                    sRect.Bottom = Height
                    sRect.Left = (AnimEditorFrame(i) - 1) * Width
                    sRect.Right = sRect.Left + Width
                    
                    dRect.Top = 0
                    dRect.Bottom = Height
                    dRect.Left = 0
                    dRect.Right = Width
                    
                    RenderTextureByRects Tex_Animation(Animationnum), sRect, dRect
                    
                    With srcRect
                        .X1 = 0
                        .X2 = frmEditor_Animation.picSprite(i).Width
                        .Y1 = 0
                        .Y2 = frmEditor_Animation.picSprite(i).Height
                    End With
                                
                    With destRECT
                        .X1 = 0
                        .X2 = frmEditor_Animation.picSprite(i).Width
                        .Y1 = 0
                        .Y2 = frmEditor_Animation.picSprite(i).Height
                    End With
                                
                    Direct3D_Device.EndScene
                    Direct3D_Device.Present srcRect, destRECT, frmEditor_Animation.picSprite(i).hwnd, ByVal (0)
                End If
            End If
        End If
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "EditorAnim_DrawAnim", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorNPC_DrawSprite()
    Dim Sprite As Long, destRECT As D3DRECT
    Dim sRect As RECT
    Dim dRect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Sprite = frmEditor_NPC.scrlSprite.Value

    If Sprite < 1 Or Sprite > NumCharacters Then
        frmEditor_NPC.picSprite.Cls
        Exit Sub
    End If

    sRect.Top = 0
    sRect.Bottom = Tex_Character(Sprite).Height / 4
    sRect.Left = (Tex_Character(Sprite).Width / 4) * 2 ' Facing down
    sRect.Right = sRect.Left + Tex_Character(Sprite).Width / 4
    dRect.Top = 0
    dRect.Bottom = Tex_Character(Sprite).Height / 4
    dRect.Left = 0
    dRect.Right = Tex_Character(Sprite).Width / 4
    
    frmEditor_NPC.picSprite.Width = Tex_Character(Sprite).Width / 4
    frmEditor_NPC.picSprite.Height = Tex_Character(Sprite).Height / 4
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Character(Sprite), sRect, dRect
    
    With destRECT
        .X1 = 0
        .X2 = frmEditor_NPC.picSprite.Width
        .Y1 = 0
        .Y2 = frmEditor_NPC.picSprite.Height
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRECT, destRECT, frmEditor_NPC.picSprite.hwnd, ByVal (0)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "EditorNpc_DrawSprite", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorResource_DrawSprite()
    Dim Sprite As Long
    Dim sRect As RECT, destRECT As D3DRECT, srcRect As D3DRECT
    Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Normal sprite
    Sprite = frmEditor_Resource.scrlNormalPic.Value

    If Sprite < 1 Or Sprite > NumResources Then
        frmEditor_Resource.picNormalPic.Cls
    Else
        sRect.Top = 0
        sRect.Bottom = Tex_Resource(Sprite).Height
        sRect.Left = 0
        sRect.Right = Tex_Resource(Sprite).Width
        dRect.Top = 0
        dRect.Bottom = Tex_Resource(Sprite).Height
        dRect.Left = 0
        dRect.Right = Tex_Resource(Sprite).Width
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
        RenderTextureByRects Tex_Resource(Sprite), sRect, dRect
        
        With srcRect
            .X1 = 0
            .X2 = Tex_Resource(Sprite).Width
            .Y1 = 0
            .Y2 = Tex_Resource(Sprite).Height
        End With
        
        With destRECT
            .X1 = 0
            .X2 = frmEditor_Resource.picNormalPic.ScaleWidth
            .Y1 = 0
            .Y2 = frmEditor_Resource.picNormalPic.ScaleHeight
        End With
                    
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRECT, frmEditor_Resource.picNormalPic.hwnd, ByVal (0)
    End If

    ' Exhausted sprite
    Sprite = frmEditor_Resource.scrlExhaustedPic.Value

    If Sprite < 1 Or Sprite > NumResources Then
        frmEditor_Resource.picExhaustedPic.Cls
    Else
        sRect.Top = 0
        sRect.Bottom = Tex_Resource(Sprite).Height
        sRect.Left = 0
        sRect.Right = Tex_Resource(Sprite).Width
        dRect.Top = 0
        dRect.Bottom = Tex_Resource(Sprite).Height
        dRect.Left = 0
        dRect.Right = Tex_Resource(Sprite).Width
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
        RenderTextureByRects Tex_Resource(Sprite), sRect, dRect
        
        With destRECT
            .X1 = 0
            .X2 = frmEditor_Resource.picExhaustedPic.ScaleWidth
            .Y1 = 0
            .Y2 = frmEditor_Resource.picExhaustedPic.ScaleHeight
        End With
        
        With srcRect
            .X1 = 0
            .X2 = Tex_Resource(Sprite).Width
            .Y1 = 0
            .Y2 = Tex_Resource(Sprite).Height
        End With
                    
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRECT, frmEditor_Resource.picExhaustedPic.hwnd, ByVal (0)
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "EditorResource_DrawSprite", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorMap_DrawMapItem()
    Dim ItemNum As Long
    Dim sRect As RECT, destRECT As D3DRECT
    Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ItemNum = Item(frmEditor_Map.scrlMapItem.Value).Pic

    If ItemNum < 1 Or ItemNum > NumItems Then
        frmEditor_Map.picMapItem.Cls
        Exit Sub
    End If

    sRect.Top = 0
    sRect.Bottom = PIC_Y
    sRect.Left = 0
    sRect.Right = PIC_X
    dRect.Top = 0
    dRect.Bottom = PIC_Y
    dRect.Left = 0
    dRect.Right = PIC_X
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Item(ItemNum), sRect, dRect
    
    With destRECT
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRECT, destRECT, frmEditor_Map.picMapItem.hwnd, ByVal (0)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "EditorMap_DrawMapItem", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorItem_DrawItem()
    Dim ItemNum As Long
    Dim sRect As RECT, destRECT As D3DRECT
    Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ItemNum = frmEditor_Item.scrlPic.Value

    If ItemNum < 1 Or ItemNum > NumItems Then
        frmEditor_Item.picItem.Cls
        Exit Sub
    End If

    ' Rect for source
    sRect.Top = 0
    sRect.Bottom = PIC_Y
    sRect.Left = 0
    sRect.Right = PIC_X
    
    ' Same for destination as source
    dRect = sRect
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Item(ItemNum), sRect, dRect
    
    With destRECT
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRECT, destRECT, frmEditor_Item.picItem.hwnd, ByVal (0)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "EditorItem_DrawItem", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorItem_DrawPaperdoll()
    Dim Sprite As Long, srcRect As D3DRECT, destRECT As D3DRECT
    Dim sRect As RECT
    Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Sprite = frmEditor_Item.scrlPaperdoll.Value

    If Sprite < 1 Or Sprite > NumPaperdolls Then
        frmEditor_Item.picPaperdoll.Cls
        Exit Sub
    End If

    ' Rect for source
    sRect.Top = 0
    sRect.Bottom = Tex_Paperdoll(Sprite).Height / 4
    sRect.Left = 0
    sRect.Right = Tex_Paperdoll(Sprite).Width / 4
    
    ' Same for destination as source
    dRect = sRect
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Paperdoll(Sprite), sRect, dRect
                    
    With destRECT
        .X1 = 0
        .X2 = Tex_Paperdoll(Sprite).Width / 4
        .Y1 = 0
        .Y2 = Tex_Paperdoll(Sprite).Height / 4
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRECT, destRECT, frmEditor_Item.picPaperdoll.hwnd, ByVal (0)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "EditorItem_DrawPaperdoll", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorSpell_DrawIcon()
    Dim IconNum As Long, destRECT As D3DRECT
    Dim sRect As RECT
    Dim dRect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IconNum = frmEditor_Spell.scrlIcon.Value
    
    If IconNum < 1 Or IconNum > NumSpellIcons Then
        frmEditor_Spell.picSprite.Cls
        Exit Sub
    End If
    
    sRect.Top = 0
    sRect.Bottom = PIC_Y
    sRect.Left = 0
    sRect.Right = PIC_X
    dRect.Top = 0
    dRect.Bottom = PIC_Y
    dRect.Left = 0
    dRect.Right = PIC_X
    
    With destRECT
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_SpellIcon(IconNum), sRect, dRect
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRECT, destRECT, frmEditor_Spell.picSprite.hwnd, ByVal (0)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "EditorSpell_DrawIcon", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorMap_DrawTileset()
    Dim Height As Long, srcRect As D3DRECT, destRECT As D3DRECT
    Dim Width As Long
    Dim Tileset As Long
    Dim sRect As RECT
    Dim dRect As RECT, scrlX As Long, scrlY As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Find tileset number
    Tileset = frmEditor_Map.scrlTileSet.Value
    
    ' Exit out if doesn't exist
    If Tileset < 0 Or Tileset > NumTileSets Then Exit Sub
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    scrlX = frmEditor_Map.scrlPictureX.Value * PIC_X
    scrlY = frmEditor_Map.scrlPictureY.Value * PIC_Y
    
    Height = Tex_Tileset(Tileset).Height - scrlY
    Width = Tex_Tileset(Tileset).Width - scrlX
    
    sRect.Left = frmEditor_Map.scrlPictureX.Value * PIC_X
    sRect.Top = frmEditor_Map.scrlPictureY.Value * PIC_Y
    sRect.Right = sRect.Left + Width
    sRect.Bottom = sRect.Top + Height
    
    dRect.Top = 0
    dRect.Bottom = Height
    dRect.Left = 0
    dRect.Right = Width
    
    RenderTextureByRects Tex_Tileset(Tileset), sRect, dRect
    
    With destRECT
        .X1 = (EditorTileX * 32) - sRect.Left
        .X2 = (EditorTileWidth * 32) + .X1
        .Y1 = (EditorTileY * 32) - sRect.Top
        .Y2 = (EditorTileHeight * 32) + .Y1
    End With
    
    DrawSelectionBox destRECT
        
    With srcRect
        .X1 = 0
        .X2 = Width
        .Y1 = 0
        .Y2 = Height
    End With
                    
    With destRECT
        .X1 = 0
        .X2 = frmEditor_Map.picBack.ScaleWidth
        .Y1 = 0
        .Y2 = frmEditor_Map.picBack.ScaleHeight
    End With
    
    ' Now render the selection tiles and we are done!
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRECT, destRECT, frmEditor_Map.picBack.hwnd, ByVal (0)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "EditorMap_DrawTileset", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub DrawSelectionBox(dRect As D3DRECT)
    Dim Width As Long, Height As Long, x As Long, y As Long
    
    Width = dRect.X2 - dRect.X1
    Height = dRect.Y2 - dRect.Y1
    x = dRect.X1
    y = dRect.Y1
    
    If Width > 6 And Height > 6 Then
        ' Draw Box 32 by 32 at graphicselx and graphicsely
        RenderTexture Tex_Selection, x, y, 1, 1, 2, 2, 2, 2, -1 ' Top left corner
        RenderTexture Tex_Selection, x + 2, y, 3, 1, Width - 4, 2, 32 - 6, 2, -1 ' Top line
        RenderTexture Tex_Selection, x + 2 + (Width - 4), y, 29, 1, 2, 2, 2, 2, -1 ' Top right corner
        RenderTexture Tex_Selection, x, y + 2, 1, 3, 2, Height - 4, 2, 32 - 6, -1 ' Left Line
        RenderTexture Tex_Selection, x + 2 + (Width - 4), y + 2, 32 - 3, 3, 2, Height - 4, 2, 32 - 6, -1 ' Right line
        RenderTexture Tex_Selection, x, y + 2 + (Height - 4), 1, 32 - 3, 2, 2, 2, 2, -1 ' Bottom left corner
        RenderTexture Tex_Selection, x + 2 + (Width - 4), y + 2 + (Height - 4), 32 - 3, 32 - 3, 2, 2, 2, 2, -1 ' Bottom right corner
        RenderTexture Tex_Selection, x + 2, y + 2 + (Height - 4), 3, 32 - 3, Width - 4, 2, 32 - 6, 2, -1 ' Bottom line
    End If
End Sub

Public Sub DrawEvents()
    Dim sRect As RECT
    Dim Width As Long, Height As Long, i As Long, x As Long, y As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Map.EventCount <= 0 Then Exit Sub
    
    For i = 1 To Map.EventCount
        If Map.events(i).PageCount <= 0 Then
                sRect.Top = 0
                sRect.Bottom = 32
                sRect.Left = 0
                sRect.Right = 32
                RenderTexture Tex_Selection, ConvertMapX(x), ConvertMapY(y), sRect.Left, sRect.Right, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
            GoTo nextevent
        End If
        
        Width = 32
        Height = 32
    
        x = Map.events(i).x * 32
        y = Map.events(i).y * 32
        x = ConvertMapX(x)
        y = ConvertMapY(y)
    
        If i > Map.EventCount Then Exit Sub
        If 1 > Map.events(i).PageCount Then Exit Sub
        
        Select Case Map.events(i).Pages(1).GraphicType
            Case 0
                sRect.Top = 0
                sRect.Bottom = 32
                sRect.Left = 0
                sRect.Right = 32
                RenderTexture Tex_Selection, x, y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
            Case 1
                If Map.events(i).Pages(1).Graphic > 0 And Map.events(i).Pages(1).Graphic <= NumCharacters Then
                    
                    sRect.Top = (Map.events(i).Pages(1).GraphicY * (Tex_Character(Map.events(i).Pages(1).Graphic).Height / 4))
                    sRect.Left = (Map.events(i).Pages(1).GraphicX * (Tex_Character(Map.events(i).Pages(1).Graphic).Width / 4))
                    sRect.Bottom = sRect.Top + 32
                    sRect.Right = sRect.Left + 32
                    RenderTexture Tex_Character(Map.events(i).Pages(1).Graphic), x, y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                    
                    sRect.Top = 0
                    sRect.Bottom = 32
                    sRect.Left = 0
                    sRect.Right = 32
                    RenderTexture Tex_Selection, x, y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                Else
                    sRect.Top = 0
                    sRect.Bottom = 32
                    sRect.Left = 0
                    sRect.Right = 32
                    RenderTexture Tex_Selection, x, y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
            Case 2
                If Map.events(i).Pages(1).Graphic > 0 And Map.events(i).Pages(1).Graphic < NumTileSets Then
                    sRect.Top = Map.events(i).Pages(1).GraphicY * 32
                    sRect.Left = Map.events(i).Pages(1).GraphicX * 32
                    sRect.Bottom = sRect.Top + 32
                    sRect.Right = sRect.Left + 32
                    RenderTexture Tex_Tileset(Map.events(i).Pages(1).Graphic), x, y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                    
                    sRect.Top = 0
                    sRect.Bottom = 32
                    sRect.Left = 0
                    sRect.Right = 32
                    RenderTexture Tex_Selection, x, y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                Else
                    sRect.Top = 0
                    sRect.Bottom = 32
                    sRect.Left = 0
                    sRect.Right = 32
                    RenderTexture Tex_Selection, x, y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
        End Select
        
nextevent:
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "DrawEvents", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorEvent_DrawGraphic()
    Dim sRect As RECT, destRECT As D3DRECT, srcRect As D3DRECT
    Dim dRect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Events.picGraphicSel.Visible Then
        Select Case frmEditor_Events.cmbGraphic.ListIndex
            Case 0
                ' None
                frmEditor_Events.picGraphicSel.Cls
                Exit Sub
            Case 1
                If frmEditor_Events.scrlGraphic.Value > 0 And frmEditor_Events.scrlGraphic.Value <= NumCharacters Then
                    If Tex_Character(frmEditor_Events.scrlGraphic.Value).Width > 793 Then
                        sRect.Left = frmEditor_Events.hScrlGraphicSel.Value
                        sRect.Right = sRect.Left + (Tex_Character(frmEditor_Events.scrlGraphic.Value).Width - sRect.Left)
                    Else
                        sRect.Left = 0
                        sRect.Right = Tex_Character(frmEditor_Events.scrlGraphic.Value).Width
                    End If
                    
                    If Tex_Character(frmEditor_Events.scrlGraphic.Value).Height > 472 Then
                        sRect.Top = frmEditor_Events.hScrlGraphicSel.Value
                        sRect.Bottom = sRect.Top + (Tex_Character(frmEditor_Events.scrlGraphic.Value).Height - sRect.Top)
                    Else
                        sRect.Top = 0
                        sRect.Bottom = Tex_Character(frmEditor_Events.scrlGraphic.Value).Height
                    End If
                    
                    With dRect
                        .Top = 0
                        .Bottom = sRect.Bottom - sRect.Top
                        .Left = 0
                        .Right = sRect.Right - sRect.Left
                    End With
                    
                    With destRECT
                        .X1 = dRect.Left
                        .X2 = dRect.Right
                        .Y1 = dRect.Top
                        .Y2 = dRect.Bottom
                    End With
                    
                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    RenderTextureByRects Tex_Character(frmEditor_Events.scrlGraphic.Value), sRect, dRect
                    
                    If (GraphicSelX2 < GraphicSelX Or GraphicSelY2 < GraphicSelY) Or (GraphicSelX2 = 0 And GraphicSelY2 = 0) Then
                        With destRECT
                            .X1 = (GraphicSelX * (Tex_Character(frmEditor_Events.scrlGraphic.Value).Width / 4)) - sRect.Left
                            .X2 = (Tex_Character(frmEditor_Events.scrlGraphic.Value).Width / 4) + .X1
                            .Y1 = (GraphicSelY * (Tex_Character(frmEditor_Events.scrlGraphic.Value).Height / 4)) - sRect.Top
                            .Y2 = (Tex_Character(frmEditor_Events.scrlGraphic.Value).Height / 4) + .Y1
                        End With

                    Else
                        With destRECT
                            .X1 = (GraphicSelX * 32) - sRect.Left
                            .X2 = ((GraphicSelX2 - GraphicSelX) * 32) + .X1
                            .Y1 = (GraphicSelY * 32) - sRect.Top
                            .Y2 = ((GraphicSelY2 - GraphicSelY) * 32) + .Y1
                        End With
                    End If
                    DrawSelectionBox destRECT
                    
                    With srcRect
                        .X1 = dRect.Left
                        .X2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .Y1 = dRect.Top
                        .Y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    With destRECT
                        .X1 = 0
                        .Y1 = 0
                        .X2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .Y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    Direct3D_Device.EndScene
                    Direct3D_Device.Present srcRect, destRECT, frmEditor_Events.picGraphicSel.hwnd, ByVal (0)
                    
                    If GraphicSelX <= 3 And GraphicSelY <= 3 Then
                    Else
                        GraphicSelX = 0
                        GraphicSelY = 0
                    End If
                Else
                    frmEditor_Events.picGraphicSel.Cls
                    Exit Sub
                End If
            Case 2
                If frmEditor_Events.scrlGraphic.Value > 0 And frmEditor_Events.scrlGraphic.Value <= NumTileSets Then
                    
                    If Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Width > 793 Then
                        sRect.Left = frmEditor_Events.hScrlGraphicSel.Value
                        sRect.Right = sRect.Left + 800
                    Else
                        sRect.Left = 0
                        sRect.Right = Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Width
                        sRect.Left = frmEditor_Events.hScrlGraphicSel.Value = 0
                    End If
                    
                    If Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Height > 472 Then
                        sRect.Top = frmEditor_Events.vScrlGraphicSel.Value
                        sRect.Bottom = sRect.Top + 512
                    Else
                        sRect.Top = 0
                        sRect.Bottom = Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Height
                        frmEditor_Events.vScrlGraphicSel.Value = 0
                    End If
                    
                    If sRect.Left = -1 Then sRect.Left = 0
                    If sRect.Top = -1 Then sRect.Top = 0
                    
                    With dRect
                        .Top = 0
                        .Bottom = sRect.Bottom - sRect.Top
                        .Left = 0
                        .Right = sRect.Right - sRect.Left
                    End With
                    
                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    RenderTextureByRects Tex_Tileset(frmEditor_Events.scrlGraphic.Value), sRect, dRect
                    
                    If (GraphicSelX2 < GraphicSelX Or GraphicSelY2 < GraphicSelY) Or (GraphicSelX2 = 0 And GraphicSelY2 = 0) Then
                        With destRECT
                            .X1 = (GraphicSelX * 32) - sRect.Left
                            .X2 = PIC_X + .X1
                            .Y1 = (GraphicSelY * 32) - sRect.Top
                            .Y2 = PIC_Y + .Y1
                        End With

                    Else
                        With destRECT
                            .X1 = (GraphicSelX * 32) - sRect.Left
                            .X2 = ((GraphicSelX2 - GraphicSelX) * 32) + .X1
                            .Y1 = (GraphicSelY * 32) - sRect.Top
                            .Y2 = ((GraphicSelY2 - GraphicSelY) * 32) + .Y1
                        End With
                    End If
                    
                    DrawSelectionBox destRECT
                    
                    With srcRect
                        .X1 = dRect.Left
                        .X2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .Y1 = dRect.Top
                        .Y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    With destRECT
                        .X1 = 0
                        .Y1 = 0
                        .X2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .Y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    Direct3D_Device.EndScene
                    Direct3D_Device.Present srcRect, destRECT, frmEditor_Events.picGraphicSel.hwnd, ByVal (0)
                Else
                    frmEditor_Events.picGraphicSel.Cls
                    Exit Sub
                End If
        End Select
    Else
        Select Case tmpEvent.Pages(curPageNum).GraphicType
            Case 0
                frmEditor_Events.picGraphic.Cls
            Case 1
                If tmpEvent.Pages(curPageNum).Graphic > 0 And tmpEvent.Pages(curPageNum).Graphic <= NumCharacters Then
                    sRect.Top = tmpEvent.Pages(curPageNum).GraphicY * (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Height / 4)
                    sRect.Left = tmpEvent.Pages(curPageNum).GraphicX * (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Width / 4)
                    sRect.Bottom = sRect.Top + (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Height / 4)
                    sRect.Right = sRect.Left + (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Width / 4)
                    With dRect
                        dRect.Top = (193 / 2) - ((sRect.Bottom - sRect.Top) / 2)
                        dRect.Bottom = dRect.Top + (sRect.Bottom - sRect.Top)
                        dRect.Left = (121 / 2) - ((sRect.Right - sRect.Left) / 2)
                        dRect.Right = dRect.Left + (sRect.Right - sRect.Left)
                    End With
                    With destRECT
                        .X1 = dRect.Left
                        .X2 = dRect.Right
                        .Y1 = dRect.Top
                        .Y2 = dRect.Bottom
                    End With
                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    RenderTextureByRects Tex_Character(frmEditor_Events.scrlGraphic.Value), sRect, dRect
                    Direct3D_Device.EndScene
                    Direct3D_Device.Present destRECT, destRECT, frmEditor_Events.picGraphic.hwnd, ByVal (0)
                End If
            Case 2
                If tmpEvent.Pages(curPageNum).Graphic > 0 And tmpEvent.Pages(curPageNum).Graphic <= NumTileSets Then
                    If tmpEvent.Pages(curPageNum).GraphicX2 = 0 Or tmpEvent.Pages(curPageNum).GraphicY2 = 0 Then
                        sRect.Top = tmpEvent.Pages(curPageNum).GraphicY * 32
                        sRect.Left = tmpEvent.Pages(curPageNum).GraphicX * 32
                        sRect.Bottom = sRect.Top + 32
                        sRect.Right = sRect.Left + 32
                        With dRect
                            dRect.Top = (193 / 2) - ((sRect.Bottom - sRect.Top) / 2)
                            dRect.Bottom = dRect.Top + (sRect.Bottom - sRect.Top)
                            dRect.Left = (120 / 2) - ((sRect.Right - sRect.Left) / 2)
                            dRect.Right = dRect.Left + (sRect.Right - sRect.Left)
                        End With
                        With destRECT
                            .X1 = dRect.Left
                            .X2 = dRect.Right
                            .Y1 = dRect.Top
                            .Y2 = dRect.Bottom
                        End With
                        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                        Direct3D_Device.BeginScene
                        RenderTextureByRects Tex_Tileset(frmEditor_Events.scrlGraphic.Value), sRect, dRect
                        Direct3D_Device.EndScene
                        Direct3D_Device.Present destRECT, destRECT, frmEditor_Events.picGraphic.hwnd, ByVal (0)
                    Else
                        sRect.Top = tmpEvent.Pages(curPageNum).GraphicY * 32
                        sRect.Left = tmpEvent.Pages(curPageNum).GraphicX * 32
                        sRect.Bottom = sRect.Top + ((tmpEvent.Pages(curPageNum).GraphicY2 - tmpEvent.Pages(curPageNum).GraphicY) * 32)
                        sRect.Right = sRect.Left + ((tmpEvent.Pages(curPageNum).GraphicX2 - tmpEvent.Pages(curPageNum).GraphicX) * 32)
                        With dRect
                            dRect.Top = (193 / 2) - ((sRect.Bottom - sRect.Top) / 2)
                            dRect.Bottom = dRect.Top + (sRect.Bottom - sRect.Top)
                            dRect.Left = (120 / 2) - ((sRect.Right - sRect.Left) / 2)
                            dRect.Right = dRect.Left + (sRect.Right - sRect.Left)
                        End With
                        With destRECT
                            .X1 = dRect.Left
                            .X2 = dRect.Right
                            .Y1 = dRect.Top
                            .Y2 = dRect.Bottom
                        End With
                        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                        Direct3D_Device.BeginScene
                        RenderTextureByRects Tex_Tileset(frmEditor_Events.scrlGraphic.Value), sRect, dRect
                        Direct3D_Device.EndScene
                        Direct3D_Device.Present destRECT, destRECT, frmEditor_Events.picGraphic.hwnd, ByVal (0)
                    End If
                End If
        End Select
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "EditorMap_DrawKey", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawEvent(id As Long)
    Dim x As Long, y As Long, Width As Long, Height As Long, sRect As RECT, dRect As RECT, Anim As Long, spritetop As Long
    
    If Map.MapEvents(id).Visible = 0 Then Exit Sub
    If InMapEditor Then Exit Sub
    
    Select Case Map.MapEvents(id).GraphicType
        Case 0
            Exit Sub
        Case 1
            If Map.MapEvents(id).GraphicNum <= 0 Or Map.MapEvents(id).GraphicNum > NumCharacters Then Exit Sub
            Width = Tex_Character(Map.MapEvents(id).GraphicNum).Width / 4
            Height = Tex_Character(Map.MapEvents(id).GraphicNum).Height / 4
            
            ' Reset frame
            If Map.MapEvents(id).Step = 3 Then
                Anim = 0
            ElseIf Map.MapEvents(id).Step = 1 Then
                Anim = 2
            End If
            
            Select Case Map.MapEvents(id).Dir
                Case DIR_UP
                    If (Map.MapEvents(id).yOffset > 8) Then Anim = Map.MapEvents(id).Step
                Case DIR_DOWN
                    If (Map.MapEvents(id).yOffset < -8) Then Anim = Map.MapEvents(id).Step
                Case DIR_LEFT
                    If (Map.MapEvents(id).xOffset > 8) Then Anim = Map.MapEvents(id).Step
                Case DIR_RIGHT
                    If (Map.MapEvents(id).xOffset < -8) Then Anim = Map.MapEvents(id).Step
            End Select
            
            ' Set the left
            Select Case Map.MapEvents(id).ShowDir
                Case DIR_UP
                    spritetop = 3
                Case DIR_RIGHT
                    spritetop = 2
                Case DIR_DOWN
                    spritetop = 0
                Case DIR_LEFT
                    spritetop = 1
            End Select
            
            If Map.MapEvents(id).WalkAnim = 1 Then Anim = 0
            
            If Map.MapEvents(id).Moving = 0 Then Anim = Map.MapEvents(id).GraphicX
            
            With sRect
                .Top = spritetop * Height
                .Bottom = .Top + Height
                .Left = Anim * Width
                .Right = .Left + Width
            End With
        
            ' Calculate the X
            x = Map.MapEvents(id).x * PIC_X + Map.MapEvents(id).xOffset - ((Width - 32) / 2)
        
            ' Is the player's height more than 32..?
            If (Height * 4) > 32 Then
                ' Create a 32 pixel offset for larger sprites
                y = Map.MapEvents(id).y * PIC_Y + Map.MapEvents(id).yOffset - ((Height) - 32)
            Else
                ' Proceed as normal
                y = Map.MapEvents(id).y * PIC_Y + Map.MapEvents(id).yOffset
            End If
        
            ' render the actual sprite
            Call DrawSprite(Map.MapEvents(id).GraphicNum, x, y, sRect)
            
        Case 2
            If Map.MapEvents(id).GraphicNum < 1 Or Map.MapEvents(id).GraphicNum > NumTileSets Then Exit Sub
            
            If Map.MapEvents(id).GraphicY2 > 0 Or Map.MapEvents(id).GraphicX2 > 0 Then
                With sRect
                    .Top = Map.MapEvents(id).GraphicY * 32
                    .Bottom = .Top + ((Map.MapEvents(id).GraphicY2 - Map.MapEvents(id).GraphicY) * 32)
                    .Left = Map.MapEvents(id).GraphicX * 32
                    .Right = .Left + ((Map.MapEvents(id).GraphicX2 - Map.MapEvents(id).GraphicX) * 32)
                End With
            Else
                With sRect
                    .Top = Map.MapEvents(id).GraphicY * 32
                    .Bottom = .Top + 32
                    .Left = Map.MapEvents(id).GraphicX * 32
                    .Right = .Left + 32
                End With
            End If
            
            x = Map.MapEvents(id).x * 32
            y = Map.MapEvents(id).y * 32
            
            x = x - ((sRect.Right - sRect.Left) / 2)
            y = y - (sRect.Bottom - sRect.Top) + 32
            
            
            If Map.MapEvents(id).GraphicY2 > 0 Then
                RenderTexture Tex_Tileset(Map.MapEvents(id).GraphicNum), ConvertMapX(Map.MapEvents(id).x * 32), ConvertMapY((Map.MapEvents(id).y - ((Map.MapEvents(id).GraphicY2 - Map.MapEvents(id).GraphicY) - 1)) * 32), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
            Else
                RenderTexture Tex_Tileset(Map.MapEvents(id).GraphicNum), ConvertMapX(Map.MapEvents(id).x * 32), ConvertMapY(Map.MapEvents(id).y * 32), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
            End If
    End Select
End Sub

Sub HandleDeviceLost()
    ' Do a loop while device is lost
    Do While Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST
        Exit Sub
    Loop
    
    UnloadTextures
    
    ' Reset the device
    Direct3D_Device.Reset Direct3D_Window
    
    DirectX_ReInit
     
    LoadTextures
End Sub

Private Function DirectX_ReInit() As Boolean
    On Error GoTo Error_Handler

    Direct3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display_Mode 'Use the current display mode that you
                                                                    'are already on. Incase you are confused, I'm
                                                                    'talking about your current screen resolution. ;)
        
    Direct3D_Window.Windowed = True 'The app will be in windowed mode.

    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_COPY 'Refresh when the monitor does.
    Direct3D_Window.BackBufferFormat = Display_Mode.Format 'Sets the format that was retrieved into the backbuffer.
    'Creates the rendering device with some useful info, along with the info
    'we've already setup for Direct3D_Window.
    'Creates the rendering device with some useful info, along with the info
    Direct3D_Window.BackBufferCount = 1 '1 backbuffer only
    Direct3D_Window.BackBufferWidth = 800 ' FrmMain.picScreen.ScaleWidth 'Match the backbuffer width with the display width
    Direct3D_Window.BackBufferHeight = 600 'frmMain.picScreen.ScaleHeight 'Match the backbuffer height with the display height
    Direct3D_Window.hDeviceWindow = frmMain.picScreen.hwnd 'Use frmMain as the device window.
    
    With Direct3D_Device
        .SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
    
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False
        
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
    End With
    
    DirectX_ReInit = True

    Exit Function
    
Error_Handler:
    MsgBox "An error occured while initializing DirectX", vbCritical
    
    DestroyGame
    
    DirectX_ReInit = False
End Function

Public Sub UpdateCamera()
    Dim offsetX As Long, offsetY As Long, StartX As Long, StartY As Long, EndX As Long, EndY As Long
    Dim centerX As Long, centerY As Long
    
    centerX = (ScreenX \ 32) / 2
    centerY = ((ScreenY \ 32) + 1) / 2
    offsetX = TempPlayer(MyIndex).xOffset + PIC_X
    offsetY = TempPlayer(MyIndex).yOffset + PIC_Y
    StartX = GetPlayerX(MyIndex) - centerX
    StartY = GetPlayerY(MyIndex) - centerY
 
    If StartX <= 0 Then
        offsetX = 0
        If StartX = 0 Then
            If TempPlayer(MyIndex).xOffset > 0 Then
                offsetX = TempPlayer(MyIndex).xOffset
            End If
        End If
        StartX = 0
    End If

    If StartY <= 0 Then
        offsetY = 0

        If StartY = 0 Then
            If TempPlayer(MyIndex).yOffset > 0 Then
                offsetY = TempPlayer(MyIndex).yOffset
            End If
        End If

        StartY = 0
    End If

    EndX = StartX + MIN_MAPX
    EndY = StartY + MIN_MAPY
    
    If GetPlayerX(MyIndex) > centerX And EndX <= Map.MaxX Then
            StartX = StartX - 1
    End If
    
    If GetPlayerY(MyIndex) > centerY And EndY <= Map.MaxY Then
            StartY = StartY - 1
    End If
    If EndX > Map.MaxX Then
        offsetX = 32
        If EndX = Map.MaxX + 1 Then
            If TempPlayer(MyIndex).xOffset < 0 Then
                offsetX = TempPlayer(MyIndex).xOffset + PIC_X
            End If
        End If
        EndX = Map.MaxX
        StartX = EndX - MIN_MAPX
    End If

    If EndY > Map.MaxY Then


        If EndY = Map.MaxY + 1 Then
            If TempPlayer(MyIndex).yOffset <= 0 Then
                offsetY = TempPlayer(MyIndex).yOffset + PIC_Y
                StartY = EndY - MIN_MAPY - 1
                'Debug.Print "1st: " & StartY
            Else
                offsetY = 0
                StartY = EndY - MIN_MAPY
                'Debug.Print "2nd: " & StartY
            End If
        Else
            offsetY = 32
            EndY = Map.MaxY
            StartY = EndY - MIN_MAPY
        End If
        EndY = Map.MaxY


    End If

    With TileView
        .Top = StartY
        .Bottom = EndY
        .Left = StartX
        .Right = EndX
    End With

    With Camera
        .Top = offsetY
        .Bottom = .Top + ScreenY
        .Left = offsetX
        .Right = .Left + ScreenX
    End With
End Sub

Public Sub InitAutotiles()
    Dim x As Long, y As Long, layerNum As Long
    
    ' Procedure used to cache autotile positions. All positioning is
    ' independant from the tileset. Calculations are convoluted and annoying.
    ' Maths is not my strong point. Luckily we're caching them so it's a one-off
    ' thing when the map is originally loaded. As such optimisation isn't an issue.
    
    ' For simplicity's sake we cache all subtile SOURCE positions in to an array.
    ' We also give letters to each subtile for easy rendering tweaks. ;]
    
    ' First, we need to re-size the array
    ReDim Autotile(0 To Map.MaxX, 0 To Map.MaxY)
    
    ' Inner tiles (Top right subtile region)
    ' NW - a
    autoInner(1).x = 32
    autoInner(1).y = 0
    
    ' NE - b
    autoInner(2).x = 48
    autoInner(2).y = 0
    
    ' SW - c
    autoInner(3).x = 32
    autoInner(3).y = 16
    
    ' SE - d
    autoInner(4).x = 48
    autoInner(4).y = 16
    
    ' Outer Tiles - NW (bottom subtile region)
    ' NW - e
    autoNW(1).x = 0
    autoNW(1).y = 32
    
    ' NE - f
    autoNW(2).x = 16
    autoNW(2).y = 32
    
    ' SW - g
    autoNW(3).x = 0
    autoNW(3).y = 48
    
    ' SE - h
    autoNW(4).x = 16
    autoNW(4).y = 48
    
    ' Outer Tiles - NE (bottom subtile region)
    ' NW - i
    autoNE(1).x = 32
    autoNE(1).y = 32
    
    ' NE - g
    autoNE(2).x = 48
    autoNE(2).y = 32
    
    ' SW - k
    autoNE(3).x = 32
    autoNE(3).y = 48
    
    ' SE - l
    autoNE(4).x = 48
    autoNE(4).y = 48
    
    ' Outer Tiles - SW (bottom subtile region)
    ' NW - m
    autoSW(1).x = 0
    autoSW(1).y = 64
    
    ' NE - n
    autoSW(2).x = 16
    autoSW(2).y = 64
    
    ' SW - o
    autoSW(3).x = 0
    autoSW(3).y = 80
    
    ' SE - p
    autoSW(4).x = 16
    autoSW(4).y = 80
    
    ' Outer Tiles - SE (bottom subtile region)
    ' NW - q
    autoSE(1).x = 32
    autoSE(1).y = 64
    
    ' NE - r
    autoSE(2).x = 48
    autoSE(2).y = 64
    
    ' SW - s
    autoSE(3).x = 32
    autoSE(3).y = 80
    
    ' SE - t
    autoSE(4).x = 48
    autoSE(4).y = 80
    
    For x = 0 To Map.MaxX
        For y = 0 To Map.MaxY
            For layerNum = 1 To MapLayer.Layer_Count - 1
                ' calculate the subtile positions and place them
                CalculateAutotile x, y, layerNum
                ' cache the rendering state of the tiles and set them
                CacheRenderState x, y, layerNum
            Next
        Next
    Next
End Sub

Public Sub CacheRenderState(ByVal x As Long, ByVal y As Long, ByVal layerNum As Long)
    Dim quarterNum As Long

    ' Exit out early
    If x < 0 Or x > Map.MaxX Or y < 0 Or y > Map.MaxY Then Exit Sub

    With Map.Tile(x, y)
        ' check if the tile can be rendered
        If .Layer(layerNum).Tileset <= 0 Or .Layer(layerNum).Tileset > NumTileSets Then
            Autotile(x, y).Layer(layerNum).RenderState = RENDER_STATE_NONE
            Exit Sub
        End If
        
        ' check if it needs to be rendered as an autotile
        If .Autotile(layerNum) = AUTOTILE_NONE Or .Autotile(layerNum) = AUTOTILE_FAKE Then
            ' default to... default
            Autotile(x, y).Layer(layerNum).RenderState = RENDER_STATE_NORMAL
        Else
            Autotile(x, y).Layer(layerNum).RenderState = RENDER_STATE_AUTOTILE
            ' cache tileset positioning
            For quarterNum = 1 To 4
                Autotile(x, y).Layer(layerNum).srcX(quarterNum) = (Map.Tile(x, y).Layer(layerNum).x * 32) + Autotile(x, y).Layer(layerNum).QuarterTile(quarterNum).x
                Autotile(x, y).Layer(layerNum).srcY(quarterNum) = (Map.Tile(x, y).Layer(layerNum).y * 32) + Autotile(x, y).Layer(layerNum).QuarterTile(quarterNum).y
            Next
        End If
    End With
End Sub

Public Sub CalculateAutotile(ByVal x As Long, ByVal y As Long, ByVal layerNum As Long)
    ' Right, so we've split the tile block in to an easy to remember
    ' collection of letters. We now need to do the calculations to find
    ' out which little lettered block needs to be rendered. We do this
    ' by reading the surrounding tiles to check for matches.
    
    ' First we check to make sure an autotile situation is actually there.
    ' Then we calculate exactly which situation has arisen.
    ' The situations are "inner", "outer", "horizontal", "vertical" and "fill".
    
    ' Exit out if we don't have an auatotile
    If Map.Tile(x, y).Autotile(layerNum) = 0 Then Exit Sub
    
    ' Okay, we have autotiling but which one?
    Select Case Map.Tile(x, y).Autotile(layerNum)
    
        ' Normal or animated - same difference
        Case AUTOTILE_NORMAL, AUTOTILE_ANIM
            ' North West Quarter
            CalculateNW_Normal layerNum, x, y
            
            ' North East Quarter
            CalculateNE_Normal layerNum, x, y
            
            ' South West Quarter
            CalculateSW_Normal layerNum, x, y
            
            ' South East Quarter
            CalculateSE_Normal layerNum, x, y
            
        ' Cliff
        Case AUTOTILE_CLIFF
            ' North West Quarter
            CalculateNW_Cliff layerNum, x, y
            
            ' North East Quarter
            CalculateNE_Cliff layerNum, x, y
            
            ' South West Quarter
            CalculateSW_Cliff layerNum, x, y
            
            ' South East Quarter
            CalculateSE_Cliff layerNum, x, y
            
        ' Waterfalls
        Case AUTOTILE_WATERFALL
            ' North West Quarter
            CalculateNW_Waterfall layerNum, x, y
            
            ' North East Quarter
            CalculateNE_Waterfall layerNum, x, y
            
            ' South West Quarter
            CalculateSW_Waterfall layerNum, x, y
            
            ' South East Quarter
            CalculateSE_Waterfall layerNum, x, y
        
        ' Anything else
        Case Else
            ' Don't need to render anything... it's fake or not an autotile
    End Select
End Sub

' Normal autotiling
Public Sub CalculateNW_Normal(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' North West
    If checkTileMatch(layerNum, x, y, x - 1, y - 1) Then tmpTile(1) = True
    
    ' North
    If checkTileMatch(layerNum, x, y, x, y - 1) Then tmpTile(2) = True
    
    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If Not tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 1, "e"
        Case AUTO_OUTER
            placeAutotile layerNum, x, y, 1, "a"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 1, "i"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 1, "m"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 1, "q"
    End Select
End Sub

Public Sub CalculateNE_Normal(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' North
    If checkTileMatch(layerNum, x, y, x, y - 1) Then tmpTile(1) = True
    
    ' North East
    If checkTileMatch(layerNum, x, y, x + 1, y - 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 2, "j"
        Case AUTO_OUTER
            placeAutotile layerNum, x, y, 2, "b"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 2, "f"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 2, "r"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 2, "n"
    End Select
End Sub

Public Sub CalculateSW_Normal(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile(1) = True
    
    ' South West
    If checkTileMatch(layerNum, x, y, x - 1, y + 1) Then tmpTile(2) = True
    
    ' South
    If checkTileMatch(layerNum, x, y, x, y + 1) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 3, "o"
        Case AUTO_OUTER
            placeAutotile layerNum, x, y, 3, "c"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 3, "s"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 3, "g"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 3, "k"
    End Select
End Sub

Public Sub CalculateSE_Normal(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' South
    If checkTileMatch(layerNum, x, y, x, y + 1) Then tmpTile(1) = True
    
    ' South East
    If checkTileMatch(layerNum, x, y, x + 1, y + 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 4, "t"
        Case AUTO_OUTER
            placeAutotile layerNum, x, y, 4, "d"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 4, "p"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 4, "l"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 4, "h"
    End Select
End Sub

' Waterfall autotiling
Public Sub CalculateNW_Waterfall(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpTile As Boolean
    
    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, x, y, 1, "i"
    Else
        ' Edge
        placeAutotile layerNum, x, y, 1, "e"
    End If
End Sub

Public Sub CalculateNE_Waterfall(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpTile As Boolean
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, x, y, 2, "f"
    Else
        ' Edge
        placeAutotile layerNum, x, y, 2, "j"
    End If
End Sub

Public Sub CalculateSW_Waterfall(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpTile As Boolean
    
    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, x, y, 3, "k"
    Else
        ' Edge
        placeAutotile layerNum, x, y, 3, "g"
    End If
End Sub

Public Sub CalculateSE_Waterfall(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpTile As Boolean
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, x, y, 4, "h"
    Else
        ' Edge
        placeAutotile layerNum, x, y, 4, "l"
    End If
End Sub

' Cliff autotiling
Public Sub CalculateNW_Cliff(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' North West
    If checkTileMatch(layerNum, x, y, x - 1, y - 1) Then tmpTile(1) = True
    
    ' North
    If checkTileMatch(layerNum, x, y, x, y - 1) Then tmpTile(2) = True
    
    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 1, "e"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 1, "i"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 1, "m"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 1, "q"
    End Select
End Sub

Public Sub CalculateNE_Cliff(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' North
    If checkTileMatch(layerNum, x, y, x, y - 1) Then tmpTile(1) = True
    
    ' North East
    If checkTileMatch(layerNum, x, y, x + 1, y - 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 2, "j"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 2, "f"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 2, "r"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 2, "n"
    End Select
End Sub

Public Sub CalculateSW_Cliff(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile(1) = True
    
    ' South West
    If checkTileMatch(layerNum, x, y, x - 1, y + 1) Then tmpTile(2) = True
    
    ' South
    If checkTileMatch(layerNum, x, y, x, y + 1) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 3, "o"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 3, "s"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 3, "g"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 3, "k"
    End Select
End Sub

Public Sub CalculateSE_Cliff(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' South
    If checkTileMatch(layerNum, x, y, x, y + 1) Then tmpTile(1) = True
    
    ' South East
    If checkTileMatch(layerNum, x, y, x + 1, y + 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation -  Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 4, "t"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 4, "p"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 4, "l"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 4, "h"
    End Select
End Sub

Public Function checkTileMatch(ByVal layerNum As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Boolean
    ' we'll exit out early if true
    checkTileMatch = True
    
    ' if it's off the map then set it as autotile and exit out early
    If X2 < 0 Or X2 > Map.MaxX Or Y2 < 0 Or Y2 > Map.MaxY Then
        checkTileMatch = True
        Exit Function
    End If
    
    ' Fakes ALWAYS return true
    If Map.Tile(X2, Y2).Autotile(layerNum) = AUTOTILE_FAKE Then
        checkTileMatch = True
        Exit Function
    End If
    
    ' check neighbour is an autotile
    If Map.Tile(X2, Y2).Autotile(layerNum) = 0 Then
        checkTileMatch = False
        Exit Function
    End If
    
    ' check we're a matching
    If Map.Tile(X1, Y1).Layer(layerNum).Tileset <> Map.Tile(X2, Y2).Layer(layerNum).Tileset Then
        checkTileMatch = False
        Exit Function
    End If
    
    ' check tiles match
    If Map.Tile(X1, Y1).Layer(layerNum).x <> Map.Tile(X2, Y2).Layer(layerNum).x Then
        checkTileMatch = False
        Exit Function
    End If
        
    If Map.Tile(X1, Y1).Layer(layerNum).y <> Map.Tile(X2, Y2).Layer(layerNum).y Then
        checkTileMatch = False
        Exit Function
    End If
End Function

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'   All of this code is for auto tiles and the math behind generating them.
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Sub placeAutotile(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long, ByVal tileQuarter As Byte, ByVal autoTileLetter As String)
    With Autotile(x, y).Layer(layerNum).QuarterTile(tileQuarter)
        Select Case autoTileLetter
            Case "a"
                .x = autoInner(1).x
                .y = autoInner(1).y
            Case "b"
                .x = autoInner(2).x
                .y = autoInner(2).y
            Case "c"
                .x = autoInner(3).x
                .y = autoInner(3).y
            Case "d"
                .x = autoInner(4).x
                .y = autoInner(4).y
            Case "e"
                .x = autoNW(1).x
                .y = autoNW(1).y
            Case "f"
                .x = autoNW(2).x
                .y = autoNW(2).y
            Case "g"
                .x = autoNW(3).x
                .y = autoNW(3).y
            Case "h"
                .x = autoNW(4).x
                .y = autoNW(4).y
            Case "i"
                .x = autoNE(1).x
                .y = autoNE(1).y
            Case "j"
                .x = autoNE(2).x
                .y = autoNE(2).y
            Case "k"
                .x = autoNE(3).x
                .y = autoNE(3).y
            Case "l"
                .x = autoNE(4).x
                .y = autoNE(4).y
            Case "m"
                .x = autoSW(1).x
                .y = autoSW(1).y
            Case "n"
                .x = autoSW(2).x
                .y = autoSW(2).y
            Case "o"
                .x = autoSW(3).x
                .y = autoSW(3).y
            Case "p"
                .x = autoSW(4).x
                .y = autoSW(4).y
            Case "q"
                .x = autoSE(1).x
                .y = autoSE(1).y
            Case "r"
                .x = autoSE(2).x
                .y = autoSE(2).y
            Case "s"
                .x = autoSE(3).x
                .y = autoSE(3).y
            Case "t"
                .x = autoSE(4).x
                .y = autoSE(4).y
        End Select
    End With
End Sub

Public Sub DrawFog()
    Dim fogNum As Long, Color As Long, x As Long, y As Long, RenderState As Long

    fogNum = CurrentFog
    If fogNum <= 0 Or fogNum > NumFogs Then Exit Sub
    Color = D3DColorRGBA(255, 255, 255, CurrentFogOpacity)

    RenderState = 0
    
    ' Render state
    Select Case RenderState
        Case 1 ' Additive
            Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
            Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        Case 2 ' Subtractive
            Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SUBTRACT
            Direct3D_Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ZERO
            Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
    End Select
    
    For x = 0 To ((Map.MaxX * 32) / 256) + 1
        For y = 0 To ((Map.MaxY * 32) / 256) + 1
            RenderTexture Tex_Fog(fogNum), ConvertMapX((x * 256) + fogOffsetX), ConvertMapY((y * 256) + fogOffsetY), 0, 0, 256, 256, 256, 256, Color
        Next
    Next
    
    ' Reset render state
    If RenderState > 0 Then
        Direct3D_Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
    End If
End Sub

Public Sub DrawTint()
    Dim Color As Long
    
    Color = D3DColorRGBA(CurrentTintR, CurrentTintG, CurrentTintB, CurrentTintA)
    
    RenderTexture Tex_White, 0, 0, 0, 0, frmMain.picScreen.ScaleWidth, frmMain.picScreen.ScaleHeight, 32, 32, Color
End Sub

Public Sub DrawWeather()
    Dim Color As Long, i As Long, SpriteLeft As Long
    
    For i = 1 To MAX_WEATHER_PARTICLES
        If WeatherParticle(i).InUse Then
            If WeatherParticle(i).Type = WEATHER_TYPE_STORM Then
                SpriteLeft = 0
            Else
                SpriteLeft = WeatherParticle(i).Type - 1
            End If
            RenderTexture Tex_Weather, ConvertMapX(WeatherParticle(i).x), ConvertMapY(WeatherParticle(i).y), SpriteLeft * 32, 0, 32, 32, 32, 32, -1
        End If
    Next
End Sub

Public Sub EditorMapProperties_DrawPanorama()
    Dim Height As Long, srcRect As D3DRECT, destRECT As D3DRECT
    Dim Width As Long
    Dim Panorama As Long
    Dim sRect As RECT
    Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Find Panorama number
    Panorama = frmEditor_MapProperties.scrlPanorama.Value
    
    ' Exit out if doesn't exist
    If Panorama < 1 Or Panorama > NumPanoramas Then
        frmEditor_MapProperties.picPanorama.Cls
        Exit Sub
    End If
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    Height = Tex_Panorama(Panorama).Height
    Width = Tex_Panorama(Panorama).Width
    
    sRect.Left = frmEditor_Map.scrlPictureX.Value * PIC_X
    sRect.Top = frmEditor_Map.scrlPictureY.Value * PIC_Y
    sRect.Right = sRect.Left + Width
    sRect.Bottom = sRect.Top + Height
    
    dRect.Top = 0
    dRect.Bottom = Height
    dRect.Left = 0
    dRect.Right = Width
    
    RenderTextureByRects Tex_Panorama(Panorama), sRect, dRect
                    
    With destRECT
        .X1 = 0
        .X2 = frmEditor_MapProperties.picPanorama.Width
        .Y1 = 0
        .Y2 = frmEditor_MapProperties.picPanorama.Height
    End With
                
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRECT, destRECT, frmEditor_MapProperties.picPanorama.hwnd, ByVal (0)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "EditorMapProperties_DrawPanorama", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
