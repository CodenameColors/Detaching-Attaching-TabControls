Public Class Form1

  'Array of forms that have been dragged off. keeps track, and allows functionaity within the code.
  Dim CreatedForms As New List(Of Form)

  Dim CreatedWindows As New List(Of TabControl)

  Dim BeginningMPos(2) As Integer
  Dim EndMPos(2) As Integer

  Dim CurrentDragTarget As TabControl
  Dim CurrentTab As TabControl

  Dim bMouseDown As Boolean

  Private Sub TabControl1_GiveFeedback(sender As System.Object, e As System.Windows.Forms.GiveFeedbackEventArgs) Handles TabControl1.GiveFeedback
    e.UseDefaultCursors = False
  End Sub

  Private Sub TabControl1_MouseEnter(sender As System.Object, e As System.EventArgs) Handles TabControl1.MouseEnter
    'check to see if the user has a beginning drag target
    If CurrentTab.TabCount >= 1 And bMouseDown Then
      CurrentDragTarget = sender
      AddTabToControl(CurrentTab, sender)

      CurrentDragTarget = New TabControl
      CurrentTab = New TabControl
      bMouseDown = False


    End If
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

    Application.DoEvents()

    'don't let the user drag out a tab if its the last one
    'UNLESS!!! the user click and drags it back in to a tab control
    If CurrentTab.TabCount = 1 Then
      Return
    End If


    If Control.MouseButtons <> MouseButtons.Left Then
      e.Action = DragAction.Cancel

      'CHECK TO SEE IF THE USER IS DRAGGING TO AN EXISTING FORM
      If CurrentDragTarget.TabCount > 0 Then

        'ADD THE TAB DOCKING LOGIC

        'the user is trying to make a new tab window. so do that
      Else

        CreatedForms.Add(New Form)

        'set the form size and position
        CreatedForms(CreatedForms.Count - 1).Size = TabControl1.Size
        CreatedForms(CreatedForms.Count - 1).StartPosition = FormStartPosition.Manual
        CreatedForms(CreatedForms.Count - 1).Location = MousePosition
        CreatedForms(CreatedForms.Count - 1).Name = "NewForm" + CreatedForms.Count.ToString()

        'create a new tab control and fill it with previous the previous tab
        CreatedWindows.Add(New TabControl)
        CreatedWindows(CreatedWindows.Count - 1).Dock = DockStyle.Fill
        CreatedWindows(CreatedWindows.Count - 1).TabPages.Add(CurrentTab.SelectedTab)

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
      End If

      'a different button was pressed so about the drag
    Else
      e.Action = DragAction.Continue
      Me.Cursor = Cursors.Help
    End If

  End Sub

  'will keep track of the users mos pos when the click within the tabs
  Private Sub TabControl1_MouseDown(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles TabControl1.MouseDown

    System.Windows.Forms.Application.DoEvents()

    If e.Button = Windows.Forms.MouseButtons.Left Then
      BeginningMPos(1) = e.X
      BeginningMPos(2) = e.Y

      CurrentTab = DirectCast(sender, System.Windows.Forms.TabControl)
    End If
    bMouseDown = True
    'Capture = True
  End Sub

  Private Sub TabControl1_MouseUp(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles TabControl1.MouseUp
    Application.DoEvents()
    If e.Button = Windows.Forms.MouseButtons.Left Then

      'check to see if the user moved their mouse
      If (e.X - BeginningMPos(1)) + (e.Y - BeginningMPos(2)) = 0 Then
        'if they haven't moved the mouse, don't apply the drag logic
        bMouseDown = False
        Return

        'user has moved mouse so apply drag logic
      Else
        TabControl1.DoDragDrop(sender, DragDropEffects.Move)
      End If
    End If
    'delete target to avoid tab attaches that shouldn't happen
    CurrentDragTarget = New TabControl

  End Sub

  Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    'Instaniate to avoid null error
    CurrentDragTarget = New TabControl
    CurrentTab = New TabControl

  End Sub

  Private Sub TabControl1_DragEnter(sender As System.Object, e As System.Windows.Forms.DragEventArgs)
    CurrentDragTarget = New TabControl
  End Sub

  Private Sub TabControl1_DragOver(sender As System.Object, e As System.Windows.Forms.DragEventArgs)

  End Sub


  Private Sub TabControl1_MouseMove(sender As System.Object, e As System.Windows.Forms.MouseEventArgs)

  End Sub

  Private Sub TabControl1_MouseMove_1(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles TabControl1.MouseMove
    For i = 0 To CreatedWindows.Count - 1
      If CreatedWindows(i).ClientRectangle.Contains(PointToClient(e.Location)) And bMouseDown Then
        Console.WriteLine("Inside and held down mouse")
      End If
    Next
  End Sub
End Class
