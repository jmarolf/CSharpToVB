Option Explicit On
Option Infer Off
Option Strict On

Imports System.Runtime.InteropServices
Imports CSharpToVBApp

Public Class AdvancedRTB
    Inherits RichTextBox

    Sub New()
        InitializeComponent()
    End Sub

    Public Event VertScrollBarRightClicked(ByVal sender As Object, ByVal loc As Point)
    Public Event HorizScrollBarRightClicked(ByVal sender As Object, ByVal loc As Point)

    Private Const OBJID_CLIENT As Long = &HFFFFFFFC 'the hwnd parm is a handle to a scroll bar control
    Private Const OBJID_HSCROLL As Long = &HFFFFFFFA 'the horizontal scroll bar of the hwnd window
    Private Const OBJID_VSCROLL As Long = &HFFFFFFFB 'the vertical scroll bar of the hwnd window
    Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
    Private components As System.ComponentModel.IContainer
    Friend WithEvents ToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents ToolStripMenuItem2 As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem3 As ToolStripMenuItem
    Private Const WM_NCRBUTTONDOWN As Integer = &HA4

    <DllImport("user32.dll", SetLastError:=True, ThrowOnUnmappableChar:=True, CharSet:=CharSet.Auto)>
    Private Shared Function SetScrollInfo(
                                        hWnd As IntPtr,
                                        <MarshalAs(UnmanagedType.I4)> nBar As SBOrientation,
                                        <MarshalAs(UnmanagedType.Struct)> ByRef lpsi As SCROLLINFO,
                                        <MarshalAs(UnmanagedType.Bool)> bRepaint As Boolean) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetScrollInfo(
                                         hWnd As IntPtr,
                                         <MarshalAs(UnmanagedType.I4)> fnBar As SBOrientation,
                                         ByRef lpsi As SCROLLINFO) As Integer
    End Function

    <DllImport("user32.dll", EntryPoint:=NameOf(GetScrollBarInfo))>
    Private Shared Function GetScrollBarInfo(ByVal hwnd As IntPtr,
                                            ByVal idObject As Integer,
                                            ByRef psbi As SCROLLBARINFO) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    Private VertRightClicked As Boolean = False
    Private sbi As SCROLLBARINFO
    Public Enum SBOrientation As Integer
        SB_HORZ = &H0
        SB_VERT = &H1
        SB_CTL = &H2
        SB_BOTH = &H3
    End Enum

    <Flags>
    Public Enum ScrollInfoMask As UInteger
        SIF_RANGE = &H1
        SIF_PAGE = &H2
        SIF_POS = &H4
        SIF_DISABLENOSCROLL = &H8
        SIF_TRACKPOS = &H10
        SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
    End Enum

    <Serializable, StructLayout(LayoutKind.Sequential)>
    Private Structure SCROLLINFO
        Public cbSize As UInteger
        <MarshalAs(UnmanagedType.U4)> Public fMask As ScrollInfoMask
        Public nMin As Integer
        Public nMax As Integer
        Public nPage As UInteger
        Public nPos As Integer
        Public nTrackPos As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure SCROLLBARINFO
        Public cbSize As Integer
        Public rcScrollBar As RECT
        Public dxyLineButton As Integer
        Public xyThumbTop As Integer
        Public xyThumbBottom As Integer
        Public reserved As Integer
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=6, ArraySubType:=UnmanagedType.U4)>
        Public rgstate() As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure RECT
        Public left, top, right, bottom As Integer
        Public Function ToRectangle() As Rectangle
            Return New Rectangle(left, top, right - left, bottom - top)
        End Function
    End Structure

    Private Sub SetScrollPos(handle As IntPtr, SB_Orientation As SBOrientation, value As Integer, v As Boolean)
        Dim scrollinfo As New SCROLLINFO With {
            .cbSize = CUInt(Marshal.SizeOf(GetType(SCROLLINFO))),
            .fMask = ScrollInfoMask.SIF_POS
        }
        SetScrollInfo(hWnd:=handle, nBar:=SB_Orientation, lpsi:=scrollinfo, bRepaint:=v)
    End Sub
    ''' <summary>
    ''' Returns the Position of the Scroll Bar Thumb
    ''' </summary>
    ''' <param name="handle"></param>
    ''' <param name="SB_Orientation"></param>
    ''' <returns></returns>
    Private Function GetScrollPos(handle As IntPtr, SB_Orientation As SBOrientation) As Integer
        Dim si As New SCROLLINFO
        With si
            .cbSize = CUInt(Marshal.SizeOf(si))
            .fMask = ScrollInfoMask.SIF_ALL
        End With
        Dim lRet As Integer = GetScrollInfo(handle, SB_Orientation, si)
        Return If(lRet <> 0, si.nPos, -1)
    End Function

    ''' <summary>
    ''' Gets and Sets the Horizontal Scroll position of the control.
    ''' </summary>
    Public Property HScrollPos() As Integer
        Get
            Return GetScrollPos(Handle, SBOrientation.SB_HORZ)

        End Get
        Set(ByVal value As Integer)
            SetScrollPos(Handle, SBOrientation.SB_HORZ, value, True)
        End Set
    End Property

    ''' <summary>
    ''' Gets and Sets the Vertical Scroll position of the control.
    ''' </summary>
    Public Property VScrollPos() As Integer
        Get
            Return GetScrollPos(Handle, SBOrientation.SB_VERT)
        End Get
        Set(ByVal value As Integer)
            SetScrollPos(Handle, SBOrientation.SB_VERT, value, True)
        End Set
    End Property

    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = WM_NCRBUTTONDOWN Then
            sbi = New SCROLLBARINFO
            sbi.cbSize = Marshal.SizeOf(sbi)
            GetScrollBarInfo(Handle, OBJID_VSCROLL, sbi)
            If sbi.rcScrollBar.ToRectangle.Contains(MousePosition) Then
                VertRightClicked = True
            Else
                sbi.cbSize = 0
                MyBase.WndProc(m)
            End If
        Else
            MyBase.WndProc(m)
            Exit Sub
        End If

        If VertRightClicked Then
            VertRightClicked = False
            ContextMenuStrip1.Show(Me, PointToClient(MousePosition))
        End If
    End Sub

    Private Sub InitializeComponent()
        components = New System.ComponentModel.Container()
        ContextMenuStrip1 = New ContextMenuStrip(components)
        ToolStripMenuItem1 = New ToolStripMenuItem()
        ToolStripMenuItem2 = New ToolStripMenuItem()
        ToolStripMenuItem3 = New ToolStripMenuItem()
        ToolStripSeparator1 = New ToolStripSeparator()
        ContextMenuStrip1.SuspendLayout()
        SuspendLayout()
        '
        'ContextMenuStrip1
        '
        ContextMenuStrip1.Items.AddRange(New ToolStripItem() {ToolStripMenuItem1, ToolStripSeparator1, ToolStripMenuItem2, ToolStripMenuItem3})
        ContextMenuStrip1.Name = NameOf(ContextMenuStrip1)
        ContextMenuStrip1.Size = New System.Drawing.Size(147, 76)
        '
        'ToolStripMenuItem1
        '
        ToolStripMenuItem1.Name = NameOf(ToolStripMenuItem1)
        ToolStripMenuItem1.Size = New System.Drawing.Size(146, 22)
        ToolStripMenuItem1.Text = "Scroll Here"
        '
        'ToolStripMenuItem2
        '
        ToolStripMenuItem2.Name = NameOf(ToolStripMenuItem2)
        ToolStripMenuItem2.Size = New System.Drawing.Size(146, 22)
        ToolStripMenuItem2.Text = "Scroll Top"
        '
        'ToolStripMenuItem3
        '
        ToolStripMenuItem3.Name = NameOf(ToolStripMenuItem3)
        ToolStripMenuItem3.Size = New System.Drawing.Size(146, 22)
        ToolStripMenuItem3.Text = "Scroll Bottom"
        '
        'ToolStripSeparator1
        '
        ToolStripSeparator1.Name = NameOf(ToolStripSeparator1)
        ToolStripSeparator1.Size = New System.Drawing.Size(143, 6)
        ContextMenuStrip1.ResumeLayout(False)
        ResumeLayout(False)
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        With sbi
            Dim SB_Size As Integer = .rcScrollBar.bottom - .rcScrollBar.top
            Dim Percent As Double = 1 - (SB_Size - PointToClient(MousePosition).Y) / SB_Size
            Dim DesiredLine As Integer = CInt(Lines.Count * Percent)
            [Select](GetFirstCharIndexFromLine(DesiredLine), 0)
            ScrollToCaret()
        End With
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        SelectionStart = 0
    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        SelectionStart = Text.Length

    End Sub
End Class
