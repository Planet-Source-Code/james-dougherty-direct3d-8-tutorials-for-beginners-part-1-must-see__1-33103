Attribute VB_Name = "modD3D"
Option Explicit

Public DX As New DirectX8
Public D3D As Direct3D8
Public D3DDevice As Direct3DDevice8
Public D3DX As D3DX8

'When do we end the loop
Public EndLoop As Boolean

'This is for the preview button
Public SelectedIndex As Integer

Public Function Initialize_Win(hWnd As Long) As Boolean
On Local Error GoTo errOut
Dim D3DPP As D3DPRESENT_PARAMETERS
Dim Mode As D3DDISPLAYMODE
 
If DX Is Nothing Then Set DX = New DirectX8
If D3DX Is Nothing Then Set D3DX = New D3DX8
Set D3D = DX.Direct3DCreate
 
D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Mode
         
D3DPP.Windowed = 1
D3DPP.SwapEffect = D3DSWAPEFFECT_DISCARD
D3DPP.BackBufferFormat = Mode.Format
D3DPP.BackBufferCount = 1
D3DPP.EnableAutoDepthStencil = 1
D3DPP.AutoDepthStencilFormat = D3DFMT_D16
D3DPP.hDeviceWindow = hWnd
 
Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DPP)

Initialize_Win = True
Exit Function
 
errOut:
 Initialize_Win = False
End Function

Public Sub Render(Optional ClearColor As Long = vbBlack)
On Local Error Resume Next
'Nope not done yet
EndLoop = False

Do
 DoEvents
 D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, ClearColor, 1, 0
 D3DDevice.BeginScene
 
 D3DDevice.EndScene
 D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
Loop Until EndLoop = True

End Sub

Public Sub Cleanup_DX()
Set D3DX = Nothing
Set D3DDevice = Nothing
Set D3D = Nothing
Set DX = Nothing
End Sub
