Public Class Form1

  'Array of forms that have been dragged off. keeps track, and allows functionaity within the code.
  Dim CreatedForms As New List(Of Form)

  Dim CreatedWindows As New List(Of TabControl)

  Dim BeginningMPos(2) As Integer
  Dim EndMPos(2) As Integer

  Dim CurrentDragTarget As TabControl
  Dim CurrentTab As TabControl

  Dim bMouseDown As Boolean

  Private Declare Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal cButtons As Integer, ByVal dwExtraInfo As IntPtr)
  Private Sub PerformMouseClick(ByVal LClick_RClick_DClick As String)
    Const MOUSEEVENTF_LEFTDOWN As Integer = 2
    Const MOUSEEVENTF_LEFTUP As Integer = 4
    Const MOUSEEVENTF_RIGHTDOWN As Integer = 6
    Const MOUSEEVENTF_RIGHTUP As Integer = 8
    If LClick_RClick_DClick = "RClick" Then
      mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, IntPtr.Zero)
      mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, IntPtr.Zero)
    ElseIf LClick_RClick_DClick = "LClick" Then
      mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, IntPtr.Zero)
      mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, IntPtr.Zero)
    ElseIf LClick_RClick_DClick = "DClick" Then
      mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, IntPtr.Zero)
      mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, IntPtr.Zero)
      mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, IntPtr.Zero)
      mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, IntPtr.Zero)
    End If
  End Sub


  Private Sub TabControl1_GiveFeedback(sender As System.Object, e As System.Windows.Forms.GiveFeedbackEventArgs) Handles TabControl1.GiveFeedback
    e.UseDefaultCursors = False
  End Sub

  Private Sub TabControl1_MouseEnter(sender As System.Object, e As System.EventArgs) Handles TabControl1.MouseEnter
    'check to see if the user has a beginning drag target
    'If CurrentTab.TabCount >= 1 And bMouseDown Then
    'CurrentDragTarget = sender
    ' AddTabToControl(CurrentTab, sender)

    'CurrentDragTarget = New TabControl
    'CurrentTab = New TabControl
    'bMouseDown = False


    'End If
  End Sub

  Private Function AddTabToControl(TabToAdd As TabControl, TabControlTarget As TabControl)

    'check to see if the target is the current tab
    If TabToAdd.Equals(TabControlTarget) Then
      Return Nothing
    End If

    Dim temp As Form = TabToAdd.FindForm()
    TabControlTarget.TabPages.Add(TabToAdd.SelectedTab)

    If TabToAdd.TabCount = 0 Then
      temp.Close()
    End If
  End Function


  Private Sub TabControl1_QueryContinueDrag(sender As System.Object, e As System.Windows.Forms.QueryContinueDragEventArgs) Handles TabControl1.QueryContinueDrag

    'Application.DoEvents()
    Console.WriteLine(Cursor.Position.X.ToString())

    If Control.MouseButtons <> MouseButtons.Left Then
      e.Action = DragAction.Cancel

      'Check to see if the mouse is inside a created tab control  when mouse up occurs
      For i = 0 To CreatedWindows.Count - 1
        If CreatedWindows(i).RectangleToScreen(CreatedWindows(i).ClientRectangle).Contains(New Point(Cursor.Position.X, Cursor.Position.Y)) Then
          'there is another tab control below the mouse. SO dock to that one.

          CreatedWindows(i).TabPages.Add(CurrentTab.SelectedTab)
          If CurrentTab.TabPages.Count = 0 Then
            CurrentTab.FindForm().Close()
            CreatedWindows.Remove(CurrentTab)
            Return
          End If
          Me.Cursor = Cursors.Default
          Return
        Else

          'Next check if the user is dragging out to create a new tab, make sure they have more than one tab, then create a new window and transfer the tab
          If Not (CreatedWindows(i).RectangleToScreen(CreatedWindows(i).ClientRectangle).Contains(New Point(Cursor.Position.X, Cursor.Position.Y))) And CurrentTab.TabPages.Count > 1 Then

            CreatedForms.Add(New Form)

            'set the form size and position
            CreatedForms(CreatedForms.Count - 1).Size = TabControl1.Size
            CreatedForms(CreatedForms.Count - 1).StartPosition = FormStartPosition.Manual
            CreatedForms(CreatedForms.Count - 1).Location = MousePosition
            CreatedForms(CreatedForms.Count - 1).Name = "NewForm" + CreatedForms.Count.ToString()

            'create a new tab control and fill it with previous the previous tab
            Dim TabTemp As New TabControl
            TabTemp.Dock = DockStyle.Fill
            TabTemp.TabPages.Add(CurrentTab.SelectedTab)

            CreatedWindows.Add(TabTemp)

            'DEBUGING ONLY. TESTING DRAG TO OTHER TABS LOGIC ERASE LATER
            'CreatedWindows(CreatedWindows.Count - 1).TabPages.Add(New TabPage)

            'set up the event delegates for the new tab/window created
            CreatedWindows(CreatedWindows.Count - 1).Name = "NewWindow" + CreatedWindows.Count.ToString()
            AddHandler CreatedWindows(CreatedWindows.Count - 1).MouseDown, AddressOf TabControl1_MouseDown
            AddHandler CreatedWindows(CreatedWindows.Count - 1).MouseUp, AddressOf TabControl1_MouseUp
            AddHandler CreatedWindows(CreatedWindows.Count - 1).GiveFeedback, AddressOf TabControl1_GiveFeedback
            AddHandler CreatedWindows(CreatedWindows.Count - 1).QueryContinueDrag, AddressOf TabControl1_QueryContinueDrag
            AddHandler CreatedWindows(CreatedWindows.Count - 1).MouseEnter, AddressOf TabControl1_MouseEnter
            AddHandler CreatedWindows(CreatedWindows.Count - 1).MouseMove, AddressOf TabControl1_MouseMove_1

            'add the new tab control to the new form
            CreatedForms(CreatedForms.Count - 1).Controls.Add(CreatedWindows(CreatedWindows.Count - 1))
            CreatedForms(CreatedForms.Count - 1).Show()

            Me.Cursor = Cursors.Default
            bMouseDown = False
            Return

          End If

        End If
      Next
    Else
      e.Action = DragAction.Continue
      Me.Cursor = Cursors.Help
    End If

  End Sub

  'will keep track of the users mos pos when the click within the tabs
  Private Sub TabControl1_MouseDown(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles TabControl1.MouseDown

    If e.Button = Windows.Forms.MouseButtons.Left Then
      BeginningMPos(1) = Cursor.Position.X
      BeginningMPos(2) = Cursor.Position.Y

      CurrentTab = DirectCast(sender, System.Windows.Forms.TabControl)
      TabControl1.DoDragDrop(sender, DragDropEffects.None)
    End If
    bMouseDown = True
    'Capture = True
  End Sub

  Private Sub TabControl1_MouseUp(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles TabControl1.MouseUp

  End Sub

  Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    'Instaniate to avoid null error
    CurrentDragTarget = New TabControl
    CurrentTab = New TabControl

    CreatedWindows.Add(TabControl1)
    CreatedForms.Add(Me)

  End Sub

  Private Sub TabControl1_MouseMove_1(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles TabControl1.MouseMove
    For i = 0 To CreatedWindows.Count - 1
      If CreatedWindows(i).ClientRectangle.Contains(PointToClient(e.Location)) And bMouseDown Then
        Console.WriteLine("Inside and held down mouse")
      End If
    Next
  End Sub

  'Checks to see if the user is moused over a tab control
  'Returns the tabcontrol if yes. ; returns nothing if no
  Private Function IsInsideTab(e As System.Windows.Forms.MouseEventArgs) As Boolean
    For i = 0 To CreatedWindows.Count - 1
      If CreatedWindows(i).ClientRectangle.Contains(PointToClient(e.Location)) And bMouseDown Then
        Return True
      End If
    Next
    Return False
  End Function

  Private Function GetTargetTab(e As System.Windows.Forms.MouseEventArgs) As TabControl
    For i = 0 To CreatedWindows.Count - 1
      If CreatedWindows(i).ClientRectangle.Contains(PointToClient(e.Location)) And bMouseDown Then
        Return CreatedWindows(i)
      End If
    Next
    Return New TabControl
  End Function

  Private Sub Form1_DragLeave(sender As System.Object, e As System.EventArgs) Handles MyBase.DragLeave

  End Sub
End Class

