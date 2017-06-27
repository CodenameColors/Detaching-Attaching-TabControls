'Antonio Morales 6/27/17
'
'TAB TESTING:
'Creation of tabs by dragging them out of their current toolbar dock
'As well as being able to dock them to already existing tabs
'When the external tabs have a close event issued put them back to the main tab bar
'Moves all the internal tab data to the target of its choice as well

Public Class Form1

  'Array of forms that have been dragged off. keeps track, and allows functionaity within the code.
  Dim CreatedForms As New List(Of Form)
  Dim CreatedWindows As New List(Of TabControl)

  Dim CurrentDragTarget As TabControl
  Dim CurrentTab As TabControl

  Private Sub TabControl1_GiveFeedback(sender As System.Object, e As System.Windows.Forms.GiveFeedbackEventArgs) Handles TabControl1.GiveFeedback
    e.UseDefaultCursors = False
  End Sub

  Private Sub TabControl1_QueryContinueDrag(sender As System.Object, e As System.Windows.Forms.QueryContinueDragEventArgs) Handles TabControl1.QueryContinueDrag

    'Application.DoEvents()
    Console.WriteLine(Cursor.Position.X.ToString())

    If Control.MouseButtons <> MouseButtons.Left Then
      e.Action = DragAction.Cancel

      'Check to see if the mouse is inside a created tab control  when mouse up occurs
      For i = 0 To CreatedWindows.Count - 1
        If CreatedWindows(i).RectangleToScreen(CreatedWindows(i).ClientRectangle).Contains(New Point(Cursor.Position.X, Cursor.Position.Y)) And Not CurrentTab.Equals(CreatedWindows(i)) Then
          'there is another tab control below the mouse. SO dock to that one
          CreatedWindows(i).TabPages.Add(CurrentTab.SelectedTab)

          'if there are no more tabs after the tab transfer close the form and remove the tab from the createdwindows array to avoid future errors
          If CurrentTab.TabPages.Count = 0 Then
            CurrentTab.FindForm().Close()
            CreatedWindows.Remove(CurrentTab)
            Me.Cursor = Cursors.Default
            Return
          End If
          Me.Cursor = Cursors.Default
          Return
        End If
        Me.Cursor = Cursors.Default
      Next

      'Next check if the user is dragging out to create a new tab, make sure they have more than one tab, then create a new window and transfer the tab
      For i = 0 To CreatedWindows.Count - 1
        If Not (CreatedWindows(i).RectangleToScreen(CreatedWindows(i).ClientRectangle).Contains(New Point(Cursor.Position.X, Cursor.Position.Y))) And CurrentTab.TabPages.Count > 1 Then

          'if valid number of tabs is found but is trying to drag onto its self then don't allow drag logic
          If Not CurrentTab.Equals(CreatedWindows(i)) Then
            Return
          End If

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
          AddHandler CreatedWindows(CreatedWindows.Count - 1).GiveFeedback, AddressOf TabControl1_GiveFeedback
          AddHandler CreatedWindows(CreatedWindows.Count - 1).QueryContinueDrag, AddressOf TabControl1_QueryContinueDrag

          'add the new tab control to the new form
          CreatedForms(CreatedForms.Count - 1).Controls.Add(CreatedWindows(CreatedWindows.Count - 1))
          CreatedForms(CreatedForms.Count - 1).Show()
          AddHandler CreatedForms(CreatedForms.Count - 1).FormClosing, AddressOf FormClose

          Me.Cursor = Cursors.Default
          Return

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

      CurrentTab = DirectCast(sender, System.Windows.Forms.TabControl)
      TabControl1.DoDragDrop(sender, DragDropEffects.None)
    End If
    'Capture = True
  End Sub

  Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    'Instaniate to avoid null error
    CurrentDragTarget = New TabControl
    CurrentTab = New TabControl

    CreatedWindows.Add(TabControl1)
    CreatedForms.Add(Me)

  End Sub

  'happens before closing a form. Moves the currents forms tabs back to the main tab control.
  Private Sub FormClose(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

    If sender.Equals(Me) Then
      Return

    Else

      For Each CurrentControl As System.Windows.Forms.TabControl In DirectCast(sender, Form).Controls
        If TryCast(CurrentControl, TabControl) Is Nothing Then
          'it failed...don't do anything
        Else
          'a tab control was found so transfer the tabs to main tab control
          For Each CurrentTab As TabPage In CurrentControl.TabPages

            If CurrentControl.TabPages.Count > 1 Then
              TabControl1.TabPages.Add(CurrentTab)
            Else
              TabControl1.TabPages.Add(CurrentTab)
              CreatedWindows.Remove(CurrentControl)
            End If
          Next
        End If
      Next
    End If

    'clean up the array to avoid null errors
    CreatedForms.Remove(DirectCast(sender, Form))

  End Sub

End Class

