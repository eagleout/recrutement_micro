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
- **Ressources** : fiches de poste détaillées avec annonces prêtes à publier, script de pré-qualification téléphonique, modèles de messages, checklist d'intégration en 15 tâches.

## Structure

```
app/
  template.html   code de l'application (HTML, CSS, JavaScript, sans dépendance)
  gen.py          injecte le contenu métier dans le template → index.html
  index.html      application générée, prête à ouvrir ou héberger
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

L'application est conçue pour tourner comme artifact Claude : les candidats et les réglages sont alors enregistrés dans une base partagée, visible uniquement des personnes ayant au moins le niveau Contributor.

Ouverte ailleurs (fichier local, GitHub Pages…), elle démarre en **mode hors connexion** : elle affiche les candidats fictifs et fonctionne entièrement, mais **rien n'est enregistré**. Pour l'héberger ailleurs avec sauvegarde, il faut brancher un stockage (base de données ou `localStorage`) à la place de l'appel `window.claude.use("db")`.

## Données personnelles

Les fiches candidats contiennent des données personnelles. La CNIL recommande de ne pas conserver une candidature non retenue plus de 2 ans après le dernier contact. Les rémunérations indiquées sont des repères Île-de-France à ajuster selon votre grille ou convention collective ; le cadre réglementaire est à valider avec votre PMI.
