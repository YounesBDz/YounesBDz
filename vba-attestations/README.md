# VBA — Fusion PT/GT/AT → TRAITE + Génération Attestations Word

Système Excel + Word automatisé :

1. **Fusionne** trois tables Excel (`PT`, `GT`, `AT`) dans une table finale `TRAITE`.
2. **Génère automatiquement** une attestation Word par employé en remplaçant des placeholders, tout en conservant la mise en page du modèle (marges, logo, header/footer, tableaux).

---

## 1. Structure des feuilles Excel

| Feuille  | Rôle                                              | Colonnes |
|----------|---------------------------------------------------|----------|
| `PT`     | Table principale (employés)                       | 18       |
| `GT`     | Table secondaire — clé en col 1                   | 18       |
| `AT`     | Table secondaire — clé en col 1                   | 18       |
| `TRAITE` | Table finale fusionnée (PT + GT + AT)             | 54       |
| `MOTIF`  | Conversion `Code → Texte` (ex. `AT → Autorisation`)| 2        |

### Relations

| De              | Vers          |
|-----------------|---------------|
| `PT` colonne 8  | `GT` colonne 1 |
| `PT` colonne 10 | `AT` colonne 1 |

### Disposition de `TRAITE`

| Plage    | Source       |
|----------|--------------|
| 1 → 18   | `PT`         |
| 19 → 36  | `GT`         |
| 37 → 54  | `AT`         |

La ligne 1 contient les **en-têtes** (utilisés ensuite comme placeholders Word).

---

## 2. Installation dans le classeur Excel

1. Ouvrir Excel → `Alt + F11` pour ouvrir l’éditeur VBA.
2. **Importer le module standard** :
   - `Fichier → Importer un fichier...` → choisir `modAttestations.bas`.
3. **Coller le code de feuille** dans `TRAITE` :
   - Dans l’explorateur de projet, double-clic sur la feuille **TRAITE**.
   - Coller le contenu de `Sheet_TRAITE.cls` (en supprimant l’en-tête `VERSION 1.0 CLASS … Attribute …` si vous copiez-collez à la main).
4. Activer la référence **Microsoft Scripting Runtime** est *optionnel* — le code utilise du late-binding (`CreateObject("Scripting.Dictionary")` et `CreateObject("Word.Application")`).
5. Enregistrer le fichier en **`.xlsm`** (classeur prenant en charge les macros).

---

## 3. Boutons à créer dans Excel

Dans l’onglet **Développeur → Insérer → Bouton (contrôle de formulaire)**, créez :

| Bouton                  | Macro affectée          |
|-------------------------|-------------------------|
| `Reconstruire TRAITE`   | `Reconstruire_TRAITE`   |
| `Générer Attestations`  | `Generate_Attestations` |

---

## 4. Fonctionnement

### Saisie ID en colonne A de `TRAITE`
Dès qu’un ID est saisi en colonne A (ligne ≥ 2), la ligne est complétée automatiquement à partir de `PT`, `GT` et `AT`.

### Bouton `Reconstruire TRAITE`
Parcourt tout `PT` ligne par ligne et reconstruit intégralement `TRAITE`.
Pour chaque ligne :
- copie `PT` (1..18),
- recherche `GT` via `PT[8] = GT[1]` → copie en (19..36),
- recherche `AT` via `PT[10] = AT[1]` → copie en (37..54).

### Bouton `Générer Attestations`
Pour chaque ligne de `TRAITE` :

1. Ouvre le **modèle Word** sélectionné.
2. Lit les **en-têtes Excel** ligne 1 et remplace, partout dans le document (corps, tableaux, headers, footers), chaque placeholder de la forme `<<NOM_DE_COLONNE>>`.
3. Cas spéciaux :
   - `<<MOTIF>>` est converti via la feuille `MOTIF` (ex. `AT → Autorisation`).
   - `<<DATE>>` : si vide → date du jour (`dd/mm/yyyy`) ; sinon formate la date Excel.
   - `<<DATE_AUTO>>` : toujours la date du jour.
4. Sauvegarde un fichier `Attestation_<NOM>_<PRENOM>.docx` dans le sous-dossier `Attestations/` à côté du classeur.

La **mise en page du modèle est préservée** : seul le texte des placeholders est remplacé.

---

## 5. Modèle Word — placeholders supportés

Tout en-tête de colonne de `TRAITE` peut être utilisé comme placeholder.
Exemples typiques :

```
<<NOM>>          <<PRENOM>>       <<CIN>>
<<MOTIF>>        <<DATE>>         <<DATE_AUTO>>
<<MATRICULE>>    <<FONCTION>>     <<DEPARTEMENT>>
```

> Astuce : tapez le placeholder **d’un seul coup** dans Word (sans correction automatique qui couperait `<<` et `>>`), sinon la recherche échoue.

---

## 6. Feuille `MOTIF`

| A (Code) | B (Texte)     |
|----------|---------------|
| AT       | Autorisation  |
| M        | Maladie       |
| CP       | Congé payé    |
| ...      | ...           |

---

## 7. Fichiers du projet

| Fichier                | Rôle                                              |
|------------------------|---------------------------------------------------|
| `modAttestations.bas`  | Module standard (toutes les macros)               |
| `Sheet_TRAITE.cls`     | Code de la feuille `TRAITE` (auto-fill à la saisie)|
| `README.md`            | Ce document                                       |

---

## 8. Macros publiques disponibles

| Macro                   | Description                                                  |
|-------------------------|--------------------------------------------------------------|
| `Reconstruire_TRAITE`   | Reconstruit toute la feuille `TRAITE` à partir de PT/GT/AT.  |
| `Generate_Attestations` | Génère un `.docx` par ligne `TRAITE`.                        |
| `FillTraiteRow(row)`    | Remplit une ligne `TRAITE` à partir de l’ID en colonne A.    |
| `ReplaceInDocument`     | Remplace un placeholder dans tout le document Word.          |
| `ColumnExists`          | Vérifie qu’une colonne existe (par son en-tête).             |
| `GetColumnIndex`        | Retourne l’index d’une colonne par son en-tête.              |
| `SafeName`              | Nettoie un nom pour en faire un nom de fichier valide.       |
| `ConvertMotif`          | Convertit un code MOTIF via la feuille `MOTIF`.              |

---

## 9. Configuration

Constantes en haut de `modAttestations.bas` (modifiables) :

```vb
Public Const SHEET_PT      As String = "PT"
Public Const SHEET_GT      As String = "GT"
Public Const SHEET_AT      As String = "AT"
Public Const SHEET_TRAITE  As String = "TRAITE"
Public Const SHEET_MOTIF   As String = "MOTIF"

Public Const COL_PT_KEY_GT As Long = 8    ' PT col 8 -> GT col 1
Public Const COL_PT_KEY_AT As Long = 10   ' PT col 10 -> AT col 1

Public Const PT_COLS As Long = 18
Public Const GT_COLS As Long = 18
Public Const AT_COLS As Long = 18

Public Const OUTPUT_SUBFOLDER As String = "Attestations"
```
