# Systeme VBA - Generation Automatique de Certificats Word

Systeme VBA Excel permettant de generer automatiquement des certificats Word personnalises pour chaque employe, avec impression automatique.

---

## Fonctionnalites

- Lecture automatique des IDs employes depuis la colonne A
- Ouverture du modele Word en arriere-plan (invisible)
- Remplacement automatique des placeholders : `<<NOM>>`, `<<PRENOM>>`, `<<CIN>>`, `<<MOTIF>>`
- Compatible avec les placeholders dans les **tableaux Word**
- Compatible avec les placeholders dans les **en-tetes et pieds de page**
- Sauvegarde automatique des certificats (.docx)
- Export PDF optionnel (macro bonus)
- Impression automatique de chaque certificat
- Barre de progression dans Excel
- Gestion des erreurs robuste
- Preservation totale de la mise en page du modele

---

## Structure du Projet

```
YounesBDz/
├── ExportCertificates.bas    # Module VBA a importer dans Excel
├── GUIDE_CERTIFICATS.md      # Ce fichier (documentation)
└── README.md                 # Profil GitHub
```

---

## Pre-requis

- Microsoft Excel (2010 ou superieur)
- Microsoft Word (2010 ou superieur)
- Windows OS

---

## Installation

### Etape 1 : Preparer la structure des dossiers

Creer les dossiers suivants sur votre machine :

```
C:\Certificats\
├── Template\
│   └── certificat.docx       <- Votre modele Word ici
└── Output\                   <- Les certificats generes seront ici
    └── PDF\                  <- (Optionnel) Pour l'export PDF
```

> **Note :** Vous pouvez modifier ces chemins dans le code VBA (constantes en haut du module).

### Etape 2 : Creer le modele Word

Creez un fichier `certificat.docx` dans `C:\Certificats\Template\` avec les placeholders suivants :

```
<<NOM>>      -> Sera remplace par le nom de l'employe
<<PRENOM>>   -> Sera remplace par le prenom de l'employe
<<CIN>>      -> Sera remplace par le numero CIN
<<MOTIF>>    -> Sera remplace par le motif du certificat
```

Exemple de contenu du modele :

```
                    CERTIFICAT DE TRAVAIL

Nous soussignes, certifions que M./Mme <<PRENOM>> <<NOM>>,
titulaire de la CIN N. <<CIN>>, a travaille au sein de notre
etablissement.

Motif : <<MOTIF>>

Fait le : _______________
Signature : _______________
```

### Etape 3 : Preparer la feuille Excel

Dans votre classeur Excel, creez une feuille nommee **`Employes`** avec la structure suivante :

| Colonne A | Colonne B | Colonne C | Colonne D | Colonne E |
|-----------|-----------|-----------|-----------|-----------|
| **ID**    | **Nom**   | **Prenom**| **CIN**   | **Motif** |
| 101       | BENALI    | Ahmed     | AB123456  | Fin de contrat |
| 102       | ALAOUI    | Fatima    | CD789012  | Demission |
| 103       | TAZI      | Mohamed   | EF345678  | Retraite  |

> **Important :** Les donnees commencent a la ligne 2. La ligne 1 est reservee aux en-tetes.

### Etape 4 : Importer le module VBA

1. Ouvrez votre classeur Excel
2. Appuyez sur `Alt + F11` pour ouvrir l'editeur VBA
3. Menu **Fichier** > **Importer un fichier...**
4. Selectionnez le fichier `ExportCertificates.bas`
5. Le module apparait dans l'arborescence du projet

> **Alternative :** Copiez-collez le contenu du fichier `.bas` dans un nouveau module (clic droit sur le projet > Inserer > Module).

### Etape 5 : Activer la reference Word (optionnel)

1. Dans l'editeur VBA, menu **Outils** > **References...**
2. Cochez **Microsoft Word XX.0 Object Library**
3. Cliquez OK

> Le code fonctionne aussi sans cette reference (Late Binding).

---

## Utilisation

### Generer les certificats (DOCX + Impression)

1. Remplissez les donnees dans la feuille `Employes`
2. Appuyez sur `Alt + F8`
3. Selectionnez `ExportCertificates`
4. Cliquez **Executer**

### Generer les certificats avec export PDF

1. Appuyez sur `Alt + F8`
2. Selectionnez `ExportCertificatesAsPDF`
3. Cliquez **Executer**

Les fichiers seront generes dans :
- DOCX : `C:\Certificats\Output\Certificat_101.docx`
- PDF : `C:\Certificats\Output\PDF\Certificat_101.pdf`

---

## Configuration

Les chemins sont definis comme constantes en haut du module VBA :

```vba
Private Const TEMPLATE_PATH As String = "C:\Certificats\Template\certificat.docx"
Private Const OUTPUT_FOLDER As String = "C:\Certificats\Output\"
Private Const SHEET_NAME As String = "Employes"
```

Modifiez ces valeurs selon votre environnement.

---

## Macros Disponibles

| Macro | Description |
|-------|-------------|
| `ExportCertificates` | Genere les certificats DOCX et les imprime |
| `ExportCertificatesAsPDF` | Genere les certificats en DOCX + PDF (sans impression) |

---

## Fonctionnement Technique

```
Excel (Feuille Employes)
    |
    |-- Lit les donnees ligne par ligne
    |
    |-- Ouvre Word en arriere-plan (invisible)
    |
    |-- Pour chaque employe :
    |   |-- Ouvre le modele certificat.docx
    |   |-- Remplace <<NOM>>, <<PRENOM>>, <<CIN>>, <<MOTIF>>
    |   |-- Cherche aussi dans les tableaux Word
    |   |-- Cherche aussi dans les en-tetes/pieds de page
    |   |-- Sauvegarde sous Certificat_[ID].docx
    |   |-- Imprime le document
    |   |-- Ferme le document
    |
    |-- Ferme Word
```

---

## Depannage

| Probleme | Solution |
|----------|----------|
| "La feuille 'Employes' n'existe pas" | Renommez votre feuille en `Employes` |
| "Modele Word introuvable" | Verifiez que le fichier existe dans le chemin configure |
| "Impossible de lancer Word" | Verifiez que Microsoft Word est installe |
| Placeholders non remplaces | Verifiez qu'ils sont ecrits exactement `<<NOM>>` (avec les chevrons doubles) |
| Erreur d'impression | Verifiez qu'une imprimante est configuree par defaut |

---

## Ameliorations Futures Possibles

- Signature numerique automatique
- QR Code sur le certificat
- Photo de l'employe
- Numerotation automatique des certificats
- Interface utilisateur (UserForm)
- Impression recto-verso
- Barre de progression visuelle
- Historique des certificats generes

---

## Licence

Libre d'utilisation et de modification.
