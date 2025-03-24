Attribute VB_Name = "mod_Center_Form"
Option Explicit
Sub Center_Form(specified_form As Form, Optional reference As Variant)
        Dim new_left As Integer
        Dim new_top As Integer
  
10    If IsMissing(reference) Then
        ' Calculate where the forms top & left locations should be for
        ' the form to be centered in the screen.
20      new_left = (Screen.Width - specified_form.Width) / 2
30      new_top = (Screen.Height - specified_form.Height) / 2

40    Else
      'Calculate the position for the form to be centered in the supplied object
50      new_left = reference.Left + (reference.Width - specified_form.Width) / 2
60      new_top = reference.Top + (reference.Height - specified_form.Height) / 2

70    End If

      ' Check if new locations will put the form off of the user's screen.
80      If (new_left < 0) Then
90        new_left = 0
100     End If

110     If (new_top < 0) Then
120       new_top = 0
130     End If

      ' Set the location of the form.
140     specified_form.Move new_left, new_top
End Sub

