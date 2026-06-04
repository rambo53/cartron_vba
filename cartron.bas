	
	' Activer le filtre automatique si ce n'est pas déjà fait
    If Not export_achat.AutoFilterMode Then export_achat.Range("A1").AutoFilter
    
    ' Filtrer la colonne 1 (colonne A) pour masquer les vides
    ' "<>" indique à Excel d'ignorer les cellules vides
	Set PlageGlobale = export_achat.Range(export_achat.Cells(1, 1), export_achat.Cells(DerniereLigne, col_to_inject_bic))
    PlageGlobale.AutoFilter Field:=col_to_inject_rib, Criteria1:="<>"
	
	' STOCKER LE RÉSULTAT FILTRÉ DANS LA VARIABLE
    ' On utilise On Error au cas où le filtre ne renverrait absolument aucune ligne
    On Error Resume Next
    Set PlageVisibles = PlageGlobale.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
	
	' TRAITEMENT SUR LA VARIABLE
    ' On vérifie d'abord que notre variable n'est pas vide
    If Not PlageVisibles Is Nothing Then
	
		' on contrôle les données de notre tableau
		
		' On génère le grpheader du xml
        
        ' La boucle For Each va s'effectuer UNIQUEMENT sur les cellules visibles
        For Each Cellule In PlageVisibles
            
            ' --- TON TRAITEMENT ICI ---
            ' Exemple : On met en majuscule le résultat visible
            
        Next Cellule

    Else
        MsgBox "Le filtre n'a retourné aucune ligne visible.", vbExclamation
    End If
	
	Application.ScreenUpdating = True
	
	
	
	
	
' =========================================================================
' 1. PROCÉDURE PRINCIPALE (LIÉE AU BOUTON)
' =========================================================================
Sub Bouton1_Cliquer()
    
	' --- CONSTANTES ---
    Const MsgId As String = "SARL CARTRO"
    Const Grpg As String  = "MIXD"
    Const Nm As String = "SARL CARTRON"
    Const PmtMtd As String = "TRF"
    Const CdPmt As String = "SEPA"
    Const AdrLine As String = "1 RUE DU GENERAL BARON FABRE 56000 VANNES"
    Const Ctry As String = "FR"
    Const IBAN As String = "FR7615589569890336821614074"
    Const BIC As String = "CMBRFR2BARK"
    Const ChrgBr As String = "SLEV"
    Const CdRgltry As String = "NNN"
	' --- CONSTANTES ---
	
	Dim new_columns As Variant
    Dim export_achat As Worksheet
    Dim export_achat_name As String
    Dim plan_comptable As Worksheet
    Dim table_correspondance As Worksheet
    Dim table_correspondance_name As String
    Dim export_achat_libelle As String
    Dim Plage_libelle As Range
    Dim row_to_start As Long
    Dim col_to_inject_key_generate As Long
    Dim col_to_inject_key_find As Long
    Dim col_to_inject_iban As Long
    Dim col_to_inject_rib As Long
    Dim col_to_inject_bic As Long
    Dim PlageRecherche As Range
    Dim Col_correspondance As String
    Dim table_correspondance_intitule As Long
    Dim table_correspondance_iban As Long
    Dim table_correspondance_rib As Long
    Dim table_correspondance_bic As Long
    Dim PlageGlobale As Range
    Dim PlageVisibles As Range
	Dim Cellule_libelle As Range
	Dim Cellule_correspondance As Range
    
	new_columns = Array("libelle_clean", "key_table", "IBAN", "RIB", "BIC")
    export_achat_libelle = "Libellé"
    row_to_start = 2
    col_to_inject_key_generate = 3
    col_to_inject_key_find = 4
    col_to_inject_iban = 5
    col_to_inject_rib = 6
    col_to_inject_bic = 7
    Col_correspondance = "correspondance_coala"
    table_correspondance_name = "table_correspondance"
    export_achat_name = "export_achat"
    table_correspondance_intitule = -6
    table_correspondance_iban = -4
    table_correspondance_rib = -3
    table_correspondance_bic = -2
    
    ' --- APPEL DES FONCTIONS POUR DÉFINIR LA PLAGE ---
    Set export_achat = ActiveSheet
    
    ' 1. Appel de la fonction letter_colonne pour obtenir la lettre de la cellule d'un nom de colonne
    Set Cellule_libelle = letter_colonne(export_achat_libelle, export_achat)
   
    ' 2. Appel de la fonction Plage_to_check pour générer la plage de données (Objets = SET)
    Set Plage_libelle = Plage_to_check(Cellule_libelle, export_achat, export_achat_name, row_to_start)
	
	' lecture de la table de l'onglet "table_correspondance"
	Set table_correspondance = Get_sheet(table_correspondance_name)
	
	' 
	Set Cellule_correspondance = letter_colonne(Col_correspondance, table_correspondance)
	' Définir la zone de recherche dans la table de correspondance
	Set PlageRecherche = Plage_to_check(Cellule_correspondance, table_correspondance, table_correspondance_name, row_to_start)
	
	create_columns new_columns, export_achat
	End
	
	'For Each NomColonne In new_columns
		' 1. On cherche le numéro de la dernière colonne utilisée sur la ligne 1
		'derniereColonne = export_achat.Cells(1, export_achat.Columns.Count).End(xlToLeft).Column
		' 2. La colonne où injecter notre titre sera la suivante
		'colonneAInjecter = derniereColonne + 1
		' 3. On inscrit le nom de l'en-tête
		'export_achat.Cells(1, colonneAInjecter).Value = NomColonne
	'Next NomColonne
	
	
	
	
	generate_matrice Plage_libelle, PlageRecherche
	
    ' --- BOUCLE DE TRAITEMENT ---
    For Each Cellule In Plage_libelle
        texte_cell = Cellule.Value

        If texte_cell <> "" Then
            If InStr(texte_cell, char_to_split) > 0 Then
                TableauMorceaux = Split(texte_cell, char_to_split)
                PremierePartie = Trim(TableauMorceaux(0))
                Cellule.Offset(0, col_to_inject_key_generate).Value = PremierePartie
                
                ' LE RAPPROCHEMENT
                Set CleTrouvee = PlageRecherche.Find(What:=PremierePartie, LookIn:=xlValues, LookAt:=xlWhole)
                
                If Not CleTrouvee Is Nothing Then
                    Cellule.Offset(0, col_to_inject_key_find).Value = CleTrouvee.Offset(0, table_correspondance_intitule).Value
                    Cellule.Offset(0, col_to_inject_iban).Value = CleTrouvee.Offset(0, table_correspondance_iban).Value
                    Cellule.Offset(0, col_to_inject_rib).NumberFormat = "@"
                    Cellule.Offset(0, col_to_inject_rib).Value = CleTrouvee.Offset(0, table_correspondance_rib).Value
                    Cellule.Offset(0, col_to_inject_bic).Value = CleTrouvee.Offset(0, table_correspondance_bic).Value
                Else
                    Cellule.Offset(0, col_to_inject_key_find).Value = val_not_in_correspondance
                End If
            End If
        End If
    Next Cellule
    
    ' Réactivation d'Excel
    'Application.Calculation = xlCalculationAutomatic
    'Application.ScreenUpdating = True
    
    MsgBox "Xml validé et généré.", vbInformation
End Sub


' =========================================================================
' 2. LES FONCTIONS EXTÉRIEURES (INDÉPENDANTES)
' =========================================================================

Function letter_colonne(label_to_find As String, sheet_to_use As Worksheet) As Range
	Dim CelluleTrouvee As Range
    ' On renvoie un Objet Range (la cellule du titre), donc on utilise "As Range" et le mot "Set"
    Set CelluleTrouvee = sheet_to_use.Rows(1).Find(What:=label_to_find, LookIn:=xlValues, LookAt:=xlWhole)
	If CelluleTrouvee Is Nothing Then
		MsgBox "Erreur critique : Le label '" & label_to_find & "' est introuvable dans l'onglet '" & sheet_to_use.Name & "'.", vbCritical
		End 
    End If
	Set letter_colonne = CelluleTrouvee
End Function


Function Plage_to_check(CelluleLabel As Range, sheet_to_use As Worksheet, table_name As String, start_row As Long) As Range
    Dim LettreColonne As String
    Dim DerniereLigne As Long
    
    ' On vérifie d'abord si la cellule titre existe
    If Not CelluleLabel Is Nothing Then
        LettreColonne = Split(CelluleLabel.Address, "$")(1)
        DerniereLigne = sheet_to_use.Cells(sheet_to_use.Rows.Count, LettreColonne).End(xlUp).Row
        
        ' On attribue l'objet final au nom de la fonction avec un SET
        Set Plage_to_check = sheet_to_use.Range(LettreColonne & start_row & ":" & LettreColonne & DerniereLigne)
    Else
        MsgBox "Le label '" & table_name & "' est introuvable sur la ligne 1.", vbCritical
		End
    End If
End Function


Function Get_sheet(table_name As String) As Worksheet
	Dim ws As Worksheet
	
	On Error Resume Next
	Set ws = Sheets(table_name)
    On Error GoTo 0
	
	If ws Is Nothing Then
        MsgBox "L'onglet '" & table_name & "' est introuvable.", vbCritical
        End
    End If
	Set Get_sheet = ws
End Function

' =========================================================================
' 3. SUBPROCESS
' =========================================================================

Sub create_columns(lst_cols As Variant, sheet_new_cols As Worksheet)
	Dim derniereColonne As Long
    Dim colonneAInjecter As Long
	For Each NomColonne In lst_cols
		' 1. On cherche le numéro de la dernière colonne utilisée sur la ligne 1
		derniereColonne = sheet_new_cols.Cells(1, sheet_new_cols.Columns.Count).End(xlToLeft).Column
		' 2. La colonne où injecter notre titre sera la suivante
		colonneAInjecter = derniereColonne + 1
		' 3. On inscrit le nom de l'en-tête
		sheet_new_cols.Cells(1, colonneAInjecter).Value = NomColonne
	Next NomColonne
End Sub


Sub generate_matrice(Plage_libelle As Range, PlageRecherche As Range)
	Const val_not_in_correspondance As String = "##"
	Const char_to_split As String = " ech "
	
	Dim texte_cell As String
	Dim Cellule As Range
	Dim TableauMorceaux() As String
	Dim PremierePartie As String
	Dim CleTrouvee As Range
	
	' Optimisation de la matrice
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
	
	For Each Cellule In Plage_libelle
        texte_cell = Cellule.Value
		
		If texte_cell <> "" Then
			If InStr(texte_cell, char_to_split) > 0 Then
				TableauMorceaux = Split(texte_cell, char_to_split)
                PremierePartie = Trim(TableauMorceaux(0))
                Cellule.Offset(0, col_to_inject_key_generate).Value = PremierePartie
				
				' LE RAPPROCHEMENT
                Set CleTrouvee = PlageRecherche.Find(What:=PremierePartie, LookIn:=xlValues, LookAt:=xlWhole)
				
				If Not CleTrouvee Is Nothing Then
				Else
                End If
			End If
		End If
		
	Next Cellule
	
    ' Réactivation d'Excel
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True	
	
End Sub
