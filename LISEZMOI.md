# Notes sur le mail : complément Outlook personnel

Ce complément ajoute un bouton **Note** sur les mails ouverts. Il ouvre un panneau où l'on écrit une note personnelle.
La note est enregistrée dans le mail lui-même, sur votre boîte Microsoft 365.
Les mails annotés reçoivent la catégorie « 📝 Note ».

## Fichiers

| Fichier | À mettre sur GitHub ? | Rôle |
|---|---|---|
| `taskpane.html`, `taskpane.css`, `taskpane.js` | oui | Le panneau de note |
| `commands.html` | oui | Fichier technique exigé par Outlook |
| `assets/` (icônes) | oui | Icônes du bouton |
| `manifest.xml` | **non**, il reste sur le PC | Déclare le complément à Outlook |

## 1. Mettre les fichiers en ligne (GitHub Pages)

1. Connectez-vous sur github.com (ou créez un compte gratuit).
2. Créez un dépôt avec le bouton **New repository** :
   - nom : `outlook-notes` ;
   - visibilité : **Public** (GitHub Pages gratuit l'exige) ;
   - cochez **Add a README file**.
3. Dans le dépôt, cliquez sur **Add file → Upload files**. Faites glisser `taskpane.html`, `taskpane.css`, `taskpane.js`, `commands.html` et le dossier `assets`, puis cliquez sur **Commit changes**.
4. Allez dans **Settings → Pages**. Choisissez **Source : Deploy from a branch**, **Branch : main / (root)**, puis **Save**.
5. Attendez 1 à 2 minutes. Vérifiez que `https://VOTRE-PSEUDO.github.io/outlook-notes/taskpane.html` s'ouvre.

## 2. Compléter le manifeste

Il faut remplacer `__BASE__` dans `manifest.xml` par `https://VOTRE-PSEUDO.github.io/outlook-notes`.
(Claude peut le faire si vous lui donnez votre pseudo GitHub.)

## 3. Installer dans le nouvel Outlook

1. Dans le nouvel Outlook, cliquez sur l'icône **Applications** dans la barre de gauche, puis sur **Ajouter des applications**.
   Autre méthode : ouvrez https://aka.ms/olksideload dans le navigateur, connecté à votre compte pro.
2. Allez dans **Mes compléments**, puis en bas de la page dans **Compléments personnalisés → Ajouter un complément personnalisé → Ajouter à partir du fichier…**
3. Choisissez `manifest.xml`, puis confirmez l'installation.

Si l'option « Ajouter à partir du fichier » n'apparaît pas, c'est que l'administrateur Microsoft 365 a bloqué les compléments personnalisés.

## Utilisation

- Ouvrez un mail. Cliquez sur **Note** dans le ruban ou dans le menu **Applications** du mail.
- Écrivez votre note : elle est enregistrée automatiquement.
- Cliquez sur la **punaise** du panneau pour qu'il reste ouvert quand vous passez d'un mail à l'autre.

## Limites

- La note fait 2 000 caractères au maximum.
- La recherche Outlook ne trouve pas le texte des notes. Pour retrouver les mails annotés, filtrez sur la catégorie « 📝 Note ».
- Si vous désinstallez le complément, les notes restent dans les mails mais deviennent invisibles. Elles réapparaissent si vous réinstallez **le même** `manifest.xml`. Ne modifiez pas la ligne `<Id>`.
