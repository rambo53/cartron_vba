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
    Const IBAN_cartron As String = "FR7615589569890336821614074"
    Const BIC_cartron As String = "CMBRFR2BARK"
    Const ChrgBr As String = "SLEV"
    Const CdRgltry As String = "NNN"
	Const directory_in_name As String = "fichier_a_traiter"
	Const directory_out_name As String = "fichier_genere"
	Const libelle_clean As String = "libelle_clean"
	Const key_table As String = "key_table"
	Const IBAN As String = "IBAN"
	Const RIB As String = "RIB"
	Const BIC As String = "BIC"
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
	Dim CheminDossier As String
	Dim CheminDossierIn As String
	Dim CheminDossierPrincipal As String
	Dim ClasseurPrincipal As Workbook
	Dim Cellule_libelle_clean As Range
    Dim Cellule_key_table As Range
	Dim Cellule_IBAN As Range
	Dim Cellule_RIB As Range
	Dim Cellule_BIC As Range
	
	new_columns = Array(libelle_clean, key_table, IBAN, RIB, BIC)
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
	
	Set ClasseurPrincipal = ThisWorkbook
	CheminDossierPrincipal = ClasseurPrincipal.Path
	
	' récupération et vérification des chemins pour les fichiers de données
	CheminDossierIn = get_path_directory(CheminDossierPrincipal, directory_in_name)
	CheminDossierOut = get_path_directory(CheminDossierPrincipal, directory_out_name)
	
	inject_data_in_worbook CheminDossierIn, export_achat_name, ClasseurPrincipal
    
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
	
	' SUB
	create_columns new_columns, export_achat
	
	Set Cellule_libelle_clean = letter_colonne(libelle_clean, export_achat)
	Set Cellule_key_table = letter_colonne(key_table, export_achat)
	Set Cellule_IBAN = letter_colonne(IBAN, export_achat)
	Set Cellule_RIB = letter_colonne(RIB, export_achat)
	Set Cellule_BIC = letter_colonne(BIC, export_achat)
	
	'SUB
	generate_matrice Plage_libelle, PlageRecherche, Cellule_libelle_clean, Cellule_key_table, Cellule_IBAN, Cellule_RIB, Cellule_BIC
    
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


Function get_path_directory(path_directory As String, path_directory_concat As String) As String
	Dim path_to_return As String
	Dim FSO As Object
	path_to_return = path_directory & "\" & path_directory_concat & "\"
	Set FSO = CreateObject("Scripting.FileSystemObject")
	If Not FSO.FolderExists(path_to_return) Then
        MsgBox "Erreur : Le dossier '" & path_to_return & "' est introuvable à l'emplacement :" & path_directory, vbCritical, "Dossier Manquant"
        End
    End If
	get_path_directory = path_to_return
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


Sub inject_data_in_worbook(path_directory As String, sheet_to_create As String, ClasseurPrincipal As Workbook)
	Dim Dossier As Object
	Dim Fichier As Object
	Dim ClasseurSource As Workbook
	Dim NomOnglet As String
	Dim FSO As Object
	
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set Dossier = FSO.GetFolder(path_directory)
	
	Application.ScreenUpdating = False
    Application.DisplayAlerts = False
	
	For Each Fichier In Dossier.Files
		If (InStr(Fichier.Name, ".xls") > 0) And (Left(Fichier.Name, 2) <> "~$") Then
			' A. Ouverture de la matrice en arrière-plan
            Set ClasseurSource = Workbooks.Open(Fichier.Path)
			' SÉCURITÉ : Si l'onglet existe déjà, on le supprime pour éviter les doublons
            On Error Resume Next
            ClasseurPrincipal.Worksheets(sheet_to_create).Delete
            On Error GoTo 0
			' B. Création du nouvel onglet à la fin
            Set NouvelOnglet = ClasseurPrincipal.Worksheets.Add(After:=ClasseurPrincipal.Sheets(ClasseurPrincipal.Sheets.Count))
            NouvelOnglet.Name = sheet_to_create
			' C. Copie des données
            ClasseurSource.Sheets(1).UsedRange.Copy Destination:=NouvelOnglet.Range("A1")
            ' D. Fermeture de la matrice
            ClasseurSource.Close SaveChanges:=False
		End If
	Next Fichier
	
	Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub


Sub generate_matrice(Plage_libelle As Range, PlageRecherche As Range, Cellule_libelle_clean As Range, Cellule_key_table As Range, Cellule_IBAN As Range, Cellule_RIB As Range, Cellule_BIC As Range)
	Const val_not_in_correspondance As String = "##"
	Const char_to_split As String = " ech "
	Const intitule As String = "Intitulé"
	Const IBAN As String = "IBAN"
	Const RIB As String = "RIB"
	Const BIC As String = "BIC"
	
	Dim texte_cell As String
	Dim Cellule As Range
	Dim TableauMorceaux() As String
	Dim PremierePartie As String
	Dim CleTrouvee As Range
	' On mémorise la feuille pour que .Cells sache où écrire précisément
	Dim wsTarget As Worksheet
	Dim wsSearch As Worksheet
	
	Set wsTarget = Plage_libelle.Worksheet
	Set wsSearch = PlageRecherche.Worksheet
	
	' Optimisation de la matrice
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
	
	For Each Cellule In Plage_libelle
        texte_cell = Cellule.Value
		
		If texte_cell <> "" Then
			If InStr(texte_cell, char_to_split) > 0 Then
				TableauMorceaux = Split(texte_cell, char_to_split)
                PremierePartie = Trim(TableauMorceaux(0))
				' On récupère le numéro de la ligne actuelle
				Dim r As Long
				r = Cellule.Row
				
				wsTarget.Cells(r, Cellule_libelle_clean.Column).Value = PremierePartie
				
				' LE RAPPROCHEMENT
                Set CleTrouvee = PlageRecherche.Find(What:=PremierePartie, LookIn:=xlValues, LookAt:=xlWhole)
				
				If Not CleTrouvee Is Nothing Then
					wsTarget.Cells(r, Cellule_key_table.Column).Value = wsSearch.Cells(CleTrouvee.Row, letter_colonne(intitule, wsSearch).Column).Value
					wsTarget.Cells(r, Cellule_IBAN.Column).Value = wsSearch.Cells(CleTrouvee.Row, letter_colonne(IBAN, wsSearch).Column).Value
                
					With wsTarget.Cells(r, Cellule_RIB.Column)
						.NumberFormat = "@"
						.Value = wsSearch.Cells(CleTrouvee.Row, letter_colonne(RIB, wsSearch).Column).Value
					End With
                
					wsTarget.Cells(r, Cellule_BIC.Column).Value = wsSearch.Cells(CleTrouvee.Row, letter_colonne(BIC, wsSearch).Column).Value
				Else
					wsTarget.Cells(r, Cellule_key_table.Column).Value = val_not_in_correspondance
                End If
			End If
		End If
		
	Next Cellule
	
    ' Réactivation d'Excel
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True	
	
End Sub
