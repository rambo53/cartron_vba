Public DateDebut As Date
Public DateFin As Date
Public Annule As Boolean

Private Sub UserForm_Initialize()
    ' On met la date du jour par défaut UNIQUEMENT à l'affichage du formulaire
    txtDateDebut.Value = Format(Now, "dd/mm/yyyy")
    txtDateFin.Value = Format(Now, "dd/mm/yyyy")
    Annule = False
End Sub

Private Sub CommandButton1_Click()
    
    If Not IsDate(txtDateDebut.Value) Or Not IsDate(txtDateFin.Value) Then
        MsgBox "Attention, le format des dates est incorrect (JJ/MM/AAAA).", vbExclamation, "Erreur de saisie"
        Exit Sub
    End If
	
	DateDebut = CDate(txtDateDebut.Value)
    DateFin = CDate(txtDateFin.Value)
	
	If DateDebut > DateFin Then
		MsgBox "Attention, la date de début est après la date de fin.", vbExclamation, "Erreur de saisie"
        Exit Sub
	End If
	
	Me.Hide
	
End Sub

' Si l'utilisateur ferme la fenêtre via la croix rouge
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Annule = True
        Me.Hide
    End If
End Sub
