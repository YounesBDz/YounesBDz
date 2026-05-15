# Certificats Word automatiques (VBA Excel)

Système VBA qui lit une liste d'employés dans Excel, ouvre un modèle Word
en arrière-plan, remplace les placeholders, sauvegarde un fichier `.docx`
par employé et l'envoie à l'imprimante. Word reste invisible et est
proprement fermé à la fin.

## 1. Structure du tableau Excel

À partir de la ligne 2, dans la feuille active :

| Colonne | Champ   | Exemple   |
| ------- | ------- | --------- |
| A       | ID      | `101`     |
| B       | NOM     | `BENALI`  |
| C       | PRENOM  | `Ahmed`   |
| D       | CIN     | `AB12345` |
| E       | MOTIF   | `Congé annuel` |

La colonne A doit être contiguë (la première cellule vide marque la fin).

## 2. Modèle Word

Créez `C:\Template\certificat.docx` avec votre mise en page (logo,
tableaux, en-têtes, etc.) et insérez où vous voulez :

```
<<NOM>>     <<PRENOM>>     <<CIN>>     <<MOTIF>>
```

Les placeholders fonctionnent partout : corps de texte, **cellules de
tableau**, en-têtes, pieds de page, zones de texte.

> Astuce : tapez chaque placeholder d'un seul trait. Si Word le découpe
> en plusieurs « runs » (par exemple à cause de l'autocorrection), Find
> peut ne pas le détecter. Solution : sélectionnez le placeholder et
> retapez-le, ou utilisez `Ctrl+H` une fois pour vous assurer qu'il est
> en un seul morceau.

## 3. Installation du module VBA

1. Ouvrez votre classeur Excel contenant la liste des employés.
2. `Alt + F11` pour ouvrir l'éditeur VBA.
3. Menu **Fichier → Importer un fichier...** et sélectionnez
   `ExportCertificates.bas`.
4. (Facultatif) Adaptez les constantes en haut du module :

   ```vba
   Private Const TEMPLATE_PATH    As String  = "C:\Template\certificat.docx"
   Private Const OUTPUT_PATH      As String  = "C:\Output\"
   Private Const PRINT_AFTER_SAVE As Boolean = True
   ```

5. Enregistrez le classeur au format `.xlsm` (macro-compatible).

## 4. Utilisation

- `Alt + F8` → choisir **ExportCertificates** → **Exécuter**.
- Ou créez un bouton sur la feuille (Insertion → Forme), clic droit →
  **Affecter une macro** → `ExportCertificates`.

À la fin, une boîte de dialogue indique le nombre de certificats générés
et le temps total. Les fichiers se trouvent dans `C:\Output\` sous la
forme :

```
Certificat_101.docx
Certificat_102.docx
...
```

## 5. Comment ça marche

- Late binding (`CreateObject("Word.Application")`) : aucune référence
  Word à activer dans l'éditeur VBA.
- `wdApp.Visible = False` + `DisplayAlerts = 0` : Word travaille
  silencieusement.
- `ReplaceEverywhere` parcourt **tous les `StoryRanges`** du document
  (corps, en-têtes, pieds de page, zones de texte, notes), donc les
  placeholders situés dans des tableaux Word ou des en-têtes sont eux
  aussi remplacés.
- `SaveAs2` avec `FileFormat:=16` (`wdFormatXMLDocument`) garantit le
  format `.docx`.
- `PrintOut Background:=False` envoie le document à l'imprimante par
  défaut et attend la fin avant de passer au suivant.
- Un gestionnaire d'erreurs ferme proprement Word même en cas de
  problème, pour éviter les processus `WINWORD.EXE` fantômes.

## 6. Personnalisations courantes

| Besoin                                | Modification                                                                 |
| ------------------------------------- | ---------------------------------------------------------------------------- |
| Ne pas imprimer                       | `PRINT_AFTER_SAVE = False`                                                   |
| Exporter aussi en PDF                 | Après `SaveAs2`, ajouter `wdDoc.ExportAsFixedFormat outFile & ".pdf", 17`    |
| Ajouter un placeholder `<<DATE>>`     | Dans la boucle : `ReplaceEverywhere wdDoc, "<<DATE>>", Format(Date,"dd/mm/yyyy")` |
| Cibler une feuille précise            | Remplacer `Set ws = ActiveSheet` par `Set ws = ThisWorkbook.Sheets("Employes")` |
| Imprimante spécifique                 | Avant la boucle : `wdApp.ActivePrinter = "Nom_Imprimante"`                   |

## 7. Dépannage

- **« Modèle Word introuvable »** : vérifiez `TEMPLATE_PATH`.
- **Le placeholder n'est pas remplacé** : il est probablement coupé en
  plusieurs morceaux dans Word. Retapez-le d'un trait et resauvegardez
  le modèle.
- **Word reste ouvert en tâche de fond** : la macro nettoie déjà ce cas
  via le gestionnaire d'erreurs ; sinon, ouvrez le Gestionnaire des
  tâches et fermez `WINWORD.EXE`.
