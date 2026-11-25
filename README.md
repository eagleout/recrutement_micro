# Recrutement micro-crèches

Outil complet de recrutement pour 2 micro-crèches : **8 postes** (par crèche : 1 éducateur·rice de jeunes enfants, 1 auxiliaire de puériculture, 2 CAP AEPE), du dépôt de candidature à l'intégration de la recrue.

Deux versions, avec les mêmes règles métier :

| Version | Fichier | Pour quoi faire |
|---|---|---|
| Application web | `app/index.html` | Pipeline par étapes (glisser-déposer), notation d'entretien en un clic, tableau de bord 3D, comparateur, intégration |
| Classeur Excel | `excel/Recrutement_MicroCreches.xlsx` | Même workflow sans macro, scorecard imprimable, dashboards 3D dynamiques |
| Classeur de démo | `excel/Recrutement_MicroCreches_DEMO.xlsx` | 22 candidats fictifs pour voir les tableaux de bord en action |
| Master prompt | `docs/Master_Prompt_Recrutement_MicroCreches.md` | Spécification complète pour régénérer ou adapter l'outil avec une IA |

## Ce que contient l'outil

- **Workflow en 8 étapes** : candidature reçue → présélection CV → entretien téléphonique → entretien + mises en situation → rencontre équipe → références & conformité → proposition → embauche. Chaque étape a un délai maximal ; au-delà, une alerte de relance apparaît.
- **Scorecard** : 12 questions par poste (expérience et mises en situation), 2 par bloc de compétences, avec réponses attendues, signaux d'alerte et relances. Note de 1 à 5, score pondéré /100 selon le poste.
  - ✅ GO à partir de 75 · 🟠 à approfondir de 60 à 74,9 · ❌ non retenu sous 60
  - ⛔ **Veto** : une note de 1 en Sécurité (Q1-Q2) ou Bientraitance (Q3-Q4) écarte le candidat, quel que soit son score.
- **Conformité** : contrôle du diplôme selon le poste, dossier obligatoire avant proposition (diplôme vérifié, attestation d'honorabilité, vaccins / aptitude, références) et alerte « bloquer l'embauche » si un document manque.
- **Cadre réglementaire 2026** suivi en continu : décret n° 2025-304 (au moins un diplômé d'État de catégorie 1, au moins 40 % de catégorie 1, direction ≥ 0,5 ETP) et attestation d'honorabilité obligatoire depuis le 1er octobre 2025.
- **Import Indeed** : relevé des candidatures depuis Gmail (dans Claude) ou depuis l'export CSV d'Indeed, avec aperçu et détection des doublons.
- **Ressources** : fiches de poste détaillées avec annonces prêtes à publier, script de pré-qualification téléphonique, modèles de messages, checklist d'intégration en 15 tâches.

## Structure

```
index.html        redirection vers app/ (GitHub Pages)
app/
  template.html   code de l'application (HTML, CSS, JavaScript, sans dépendance obligatoire)
  gen.py          injecte le contenu métier dans le template → index.html
  index.html      application générée, prête à ouvrir ou héberger
  config.js       connexion Supabase (vide = stockage dans le navigateur)
supabase/
  schema.sql      tables, règles d'accès et temps réel pour le mode partagé
excel/
  content.py      contenu métier unique : fiches, 36 questions, pondérations, étapes, modèles
  build.py        génère les classeurs Excel (openpyxl)
  *.xlsx          classeurs générés
docs/
  Master_Prompt_Recrutement_MicroCreches.md
```

Le contenu métier vit à un seul endroit (`excel/content.py`) : le modifier puis relancer les deux générateurs met à jour l'application et les classeurs.

## Régénérer

```bash
pip install openpyxl
python excel/build.py excel/Recrutement_MicroCreches.xlsx          # classeur vierge (1 ligne d'exemple)
python excel/build.py excel/Recrutement_MicroCreches_DEMO.xlsx demo # classeur de démo
python app/gen.py                                                   # application
```

Les classeurs sont écrits avec leurs formules ; Excel recalcule tout à l'ouverture.

## Enregistrement des données de l'application

L'application choisit son stockage toute seule :

| Où elle tourne | Stockage | Partage |
|---|---|---|
| Dans Claude (artifact) | Base de données de l'artifact | Personnes ayant au moins le niveau Contributor |
| Hébergée, `app/config.js` rempli | Supabase (base PostgreSQL gratuite) | Tous les membres autorisés, sur tous leurs appareils, en temps réel |
| Hébergée, `app/config.js` vide | Navigateur (`localStorage`) | Un seul navigateur : faites des sauvegardes JSON |

Dans **Réglages › Stockage et sauvegarde** : export CSV, **sauvegarde complète JSON** et **import** d'une sauvegarde. Une sauvegarde faite dans un mode s'importe dans un autre (ex. du navigateur vers Supabase).

## Importer les candidatures Indeed

Indeed n'ouvre son API de candidatures (Indeed Apply, Job Sync) qu'aux éditeurs de logiciels de recrutement agréés : une connexion directe n'est pas possible pour un employeur. L'onglet **Candidats › Importer depuis Indeed** propose deux voies, avec un aperçu à valider avant toute création :

| Voie | Où | Comment |
|---|---|---|
| **Gmail** | Version dans Claude | L'application cherche les e-mails Indeed dans le Gmail connecté à Claude, Claude ne garde que les notifications de candidature (pas les alertes emploi) et en extrait nom, poste, crèche, date et profil. Recherche Gmail modifiable, e-mails déjà importés ignorés. |
| **Fichier CSV** | Partout | Export « Exporter les candidats » d'Indeed employeur (colonnes en français ou en anglais, séparateur `;` ou `,`). Indeed réserve cet export aux comptes qui atteignent un seuil de dépenses sur 12 mois (4 000 € en France) avec double authentification. |

Le poste (EJE, AP, CAP AEPE) et la crèche sont déduits de l'intitulé de l'offre et restent modifiables dans l'aperçu ; les doublons (même e-mail Indeed, même adresse ou même nom sur le même poste) sont signalés et décochés. L'onglet **Publier une offre** copie les annonces prêtes à coller dans Indeed.

Pour le relevé Gmail, le Gmail connecté à Claude doit recevoir les notifications du compte Indeed employeur (directement ou par une règle de transfert).

## Héberger l'application

### 1. Mettre en ligne avec GitHub Pages (5 minutes)

1. Sur GitHub : **Settings › Pages › Build and deployment › Deploy from a branch**, branche `main`, dossier `/ (root)`, **Save**.
2. Après une minute, l'application est en ligne à `https://<compte>.github.io/recrutement_micro/` (la page racine redirige vers `app/`).

Le site ne contient que le code de l'application et les candidats fictifs d'exemple : aucune donnée réelle n'est publiée. Sans l'étape 2, chaque navigateur garde ses propres données.

### 2. Partager les données avec Supabase (15 minutes, gratuit)

1. Créez un projet sur [supabase.com](https://supabase.com) (région Europe, par exemple Paris ou Francfort).
2. **SQL Editor** : collez le contenu de `supabase/schema.sql`, remplacez `votre.email@exemple.fr` par votre adresse, ajoutez une ligne `insert` par membre de l'équipe, puis **Run**.
3. **Authentication › URL Configuration** : *Site URL* = l'adresse GitHub Pages de l'application (`https://<compte>.github.io/recrutement_micro/app/`), et ajoutez-la aussi dans *Redirect URLs*.
4. **Project Settings › API** : copiez *Project URL* et la clé *anon public* dans `app/config.js`, puis poussez le fichier.
5. Ouvrez l'application : saisissez votre e-mail, cliquez sur le lien reçu. Si vous aviez des données dans le navigateur, importez votre sauvegarde JSON depuis Réglages.

Sécurité : la clé *anon* est publique par conception. Les données sont protégées par les règles d'accès (RLS) de `schema.sql` : seule une personne connectée **et** inscrite dans la table `membres` peut lire ou modifier les candidats. Ne mettez jamais la clé *service_role* dans `config.js`. Pour retirer l'accès à quelqu'un, supprimez sa ligne dans `membres`.

## Données personnelles

Les fiches candidats contiennent des données personnelles. La CNIL recommande de ne pas conserver une candidature non retenue plus de 2 ans après le dernier contact. Les rémunérations indiquées sont des repères Île-de-France à ajuster selon votre grille ou convention collective ; le cadre réglementaire est à valider avec votre PMI.
