# MASTER PROMPT — Workflow Excel de recrutement pour micro-crèches

> À coller tel quel dans Claude (ou tout assistant capable de générer un .xlsx). Remplacez d'abord les **[VARIABLES]**.

---

## 0. Variables à personnaliser

| Variable | Valeur |
|---|---|
| [CRÈCHE_1] / [VILLE_1] | ex. Micro-crèche Les Petitous Noisy |
| [CRÈCHE_2] / [VILLE_2] | ex. Micro-crèche Les Petitous Drancy |
| [DATE_PRISE_DE_POSTE] | ex. 01/02/2027 |
| [POSTES_PAR_CRÈCHE] | 1 EJE · 1 auxiliaire de puériculture (AP) · 2 CAP AEPE |
| [FOURCHETTES_SALAIRE] | EJE 2 200–2 800 € · AP 1 850–2 500 € · CAP AEPE SMIC–2 050 € (brut/mois) |
| [DIRECTION] | Nom de la direction de chaque crèche + ETP |

---

## 1. Rôle

Tu es à la fois **DRH spécialisé petite enfance** (EAJE, micro-crèches, cadre PMI/CAF) et **architecte Excel senior** (modélisation, dashboards, UX de classeur). Tu livres un fichier **.xlsx complet, sans macro, immédiatement utilisable**, pas un plan ni des explications.

## 2. Contexte

Je suis gestionnaire et je reprends 2 micro-crèches ([CRÈCHE_1], [CRÈCHE_2]), 12 berceaux chacune. Je dois recruter **8 personnes** : par crèche 1 EJE, 1 AP et 2 CAP AEPE. Prise de poste visée : [DATE_PRISE_DE_POSTE].

Contraintes réglementaires à intégrer dans la logique du classeur (et à rappeler sur l'accueil) :
- **Décret n° 2025-304 du 1er avril 2025** (dispositions micro-crèches applicables au **1er septembre 2026**) : au moins 1 professionnel diplômé d'État de catégorie 1 (AP, EJE, infirmier, psychomotricien) ; les catégories 1 représentent **≥ 40 % de l'effectif** ; fonction de direction **≥ 0,5 ETP**, une même personne ne dirigeant pas plus de 2 structures.
- **Attestation d'honorabilité** obligatoire depuis le **1er octobre 2025** pour tous les professionnels de crèche : l'employeur la récupère, la vérifie et la conserve, à l'embauche puis périodiquement (plateforme honorabilite.social.gouv.fr).
- Aucune « journée d'essai » non rémunérée : l'étape terrain est une **visite + rencontre d'équipe sans prise en charge d'enfants**.

## 3. Livrable attendu

Un classeur Excel de **12 onglets visibles + 1 masqué**, dans cet ordre :

1. **Accueil** — titre, tuiles de navigation cliquables (HYPERLINK vers chaque onglet), tableau des 8 étapes du workflow (délai max, responsable, livrable), « Démarrer en 5 minutes », règles de décision, cadre réglementaire avec sources, légende des couleurs.
2. **Dashboard** — filtre « Crèche » (Toutes / Crèche 1 / Crèche 2) qui pilote TOUT :
   - 6 tuiles KPI : postes pourvus (x / 8 + % couverture), candidats (dont en cours), entretiens évalués (+ score moyen), taux de GO, délai moyen candidature → embauche, alertes actives ;
   - **graphiques 3D** : entonnoir par étape (histogramme 3D), à pourvoir / pourvus / candidats par poste (histogramme 3D groupé), candidatures et embauches par source (barres 3D horizontales), comparatif Crèche 1 vs Crèche 2 (histogramme 3D) + un **radar** des moyennes par bloc de compétences et par poste ;
   - tableaux sources des graphiques visibles sous les graphiques ;
   - bloc **Conformité décret 2025-304** par crèche : % cat. 1 prévu, % cat. 1 réel (embauchés + effectif déjà en poste), ≥ 1 diplômé cat. 1, direction ≥ 0,5 ETP, statut ✅ / ⚠️ ;
   - tableau des échéances des 8 postes (jours restants, statut, recrue, alerte).
3. **Plan** — 8 lignes (P01 à P08) : crèche, poste, n° de poste, catégorie, contrat ▼, heures, dates (ouverture, prise de poste), compteurs calculés depuis le Pipeline (candidats, en cours, évalués, embauchés), score de la recrue ou meilleur score en cours, pourvu (1/0), statut (✅ Pourvu / 🔄 En cours / 🟡 À traiter / ⏳ À lancer), nom de la recrue, jours restants, alerte d'échéance (🔴 dépassée / 🟠 < 30 j / 🟢).
4. **Pipeline** — 1 ligne = 1 candidature, 150 lignes pré-numérotées (ID **statiques** C001…C150, pour que le tri ne casse aucun lien). Groupes de colonnes colorés : Identité · Candidature (crèche ▼, poste ▼, source ▼, date, diplôme ▼ + contrôle automatique diplôme/poste, expérience, disponibilité, prétentions) · Suivi (étape ▼, statut ▼, dernière action, prochaine action, note tél. /5) · Évaluation (score /100 et verdict remontés de la grille) · Conformité (diplôme vérifié, attestation d'honorabilité, vaccins/aptitude, références → « ✅ Complet » ou « ⏳ n pièce(s) manquante(s) ») · Embauche (date, délai) · **Alerte** calculée :
   - « ⚠️ Dossier incomplet — bloquer l'embauche » si étape ≥ Proposition et dossier non complet ;
   - « ⚠️ Verdict négatif — clôturer » si statut En cours et verdict Non retenu ;
   - « 🔔 Relancer : n j sans action » si le délai max de l'étape est dépassé.
   Colonnes de calcul internes masquées (nom complet, n° d'étape, libellé de liste, clés de classement, clé d'embauche). Volets figés, filtres automatiques.
5. **Scorecard** (le cœur, interactive et imprimable A4 paysage) :
   - liste déroulante du candidat (liste dynamique sans lignes vides, via plage nommée OFFSET) ; à défaut, liste « poste » pour imprimer une **grille vierge** ;
   - carte d'identité (crèche, poste, diplôme + conformité, expérience, source, étape, statut, évaluateur, date, disponibilité, prétentions, dossier) ;
   - grande tuile score /100, rang parmi les candidats de la même crèche et du même poste, nombre de notes saisies, bandeau **VERDICT** coloré ;
   - les **12 questions du poste** affichées automatiquement : bloc, type, question, réponse attendue ✅, signaux d'alerte 🚩, note, étoiles ★★★★☆ ;
   - synthèse par bloc (poids du poste, score % avec barre de données) + **radar candidat vs profil cible** ;
   - lien « ✏️ Saisir / modifier ses notes ➜ » qui ouvre directement la bonne ligne de la grille ;
   - échelle de notation et règles de décision en pied de page.
6. **Grille_Notes** — 1 ligne par ID : évaluateur, date, **Q1 à Q12 notées de 1 à 5** (validation entière 1-5 + échelle de couleurs), moyennes par bloc, score pondéré /100, saisie x/12, VETO, verdict, points forts, points de vigilance.
7. **Comparateur** — filtres crèche (dont « Toutes ») + poste → **top 10 classé par score** avec détail des 6 blocs, étape, statut, dossier, disponibilité, prétentions ; phrase « 🏆 Meilleur profil » ; histogramme 3D des scores.
8. **Onboarding** — 15 tâches (avant J1, J1, semaine 1-2, J+30, fin d'essai : contrat, DPAE, attestation d'honorabilité archivée avec date de renouvellement, diplôme, mutuelle, visite d'information et de prévention, protocoles signés, tutorat, rapport d'étonnement, bilan d'essai, plan de formation) × 8 colonnes (une par poste du Plan, nom de la recrue affiché automatiquement), statuts ▼ et % d'avancement.
9. **Fiches_Poste** — 3 cartes côte à côte (EJE, AP, CAP AEPE) : catégorie réglementaire, diplôme requis, rattachement, raison d'être, missions principales, compétences techniques, savoir-être, expérience, conditions de travail, rémunération indicative [FOURCHETTES_SALAIRE], évolution, prérequis à l'embauche, **annonce prête à publier**.
10. **Questions** — banque de **36 questions** (12 par poste, 2 par bloc) : clé poste|Qn (masquée), poste, n°, bloc, type (Expérience / Mise en situation / Connaissances / Motivation / Projet pédagogique / Animation d'équipe), question, réponse attendue, signaux d'alerte, relance.
11. **Scripts_Modeles** — script de pré-qualification téléphonique (8 questions : diplôme, expérience, horaires tournants, date de démarrage, trajet, prétentions, motivation, accord attestation d'honorabilité ; colonne « éliminatoire ») + 4 modèles de messages (convocation, demande de pièces, refus bienveillant, proposition d'embauche).
12. **Parametres** — crèches (nom, ville, capacité, dates, direction, ETP, effectif cat. 1 / cat. 2 déjà en poste), postes, étapes + délais max, **blocs & pondérations par poste** (contrôle total = 100 %), seuils, échelle de notation, toutes les listes déroulantes (plages nommées).
13. **Calculs** (masqué) — liste compacte des candidats pour la liste déroulante, critère du filtre crèche, données du radar de la Scorecard.

## 4. Logique de la scorecard (à respecter exactement)

**6 blocs, 2 questions chacun :**

| Bloc | EJE | AP | CAP AEPE | Éliminatoire |
|---|---|---|---|---|
| B1 Sécurité & santé de l'enfant | 15 % | 25 % | 20 % | Oui |
| B2 Bientraitance & posture professionnelle | 20 % | 20 % | 25 % | Oui |
| B3 Éveil & développement | 25 % | 15 % | 15 % | Non |
| B4 Relation aux familles | 15 % | 15 % | 10 % | Non |
| B5 Équipe & communication (EJE : animation d'équipe) | 15 % | 15 % | 15 % | Non |
| B6 Expérience, motivation & adéquation | 10 % | 10 % | 15 % | Non |

- Note par question : **1 Insuffisant · 2 Fragile · 3 Conforme · 4 Solide · 5 Excellent**.
- Moyenne de bloc = moyenne des notes saisies du bloc ; score = Σ(poids × moyenne) / Σ(poids des blocs notés) / 5 × 100 → l'évaluation partielle reste exploitable et est signalée « (évaluation partielle) ».
- Verdict : **✅ GO ≥ 75** · **🟠 À approfondir 60–74,9** · **❌ Non retenu < 60** · **⛔ VETO** si une note de 1 en B1 ou B2, quel que soit le score. Références « Négatif » → « ❌ Références négatives ».
- Seuils et poids modifiables dans Paramètres, jamais en dur dans les formules.

**Questions :** pour chaque poste, rédige des mises en situation concrètes de micro-crèche (chute avec choc à la tête, fièvre à 39 °C, réaction allergique avec PAI, inconnu venant chercher un enfant, étouffement, collègue qui force à manger ou menace de punition, morsures et parent en colère, familiarisation d'un bébé, parent qui refuse la sieste, absence d'une collègue à l'ouverture et taux d'encadrement, conflit d'équipe, repérage d'un trouble du développement sans poser de diagnostic…) et des questions d'expérience comportementales (« Racontez une fois où… »). Chaque question a des indicateurs de bonne réponse et des signaux d'alerte observables.

## 5. Workflow (8 étapes, délai max avant alerte)

1. Candidature reçue (2 j) → 2. Présélection CV (2 j) → 3. Entretien téléphonique (3 j) → 4. Entretien + mises en situation (5 j) → 5. Rencontre équipe / visite (5 j) → 6. Références & conformité (7 j) → 7. Proposition envoyée (5 j) → 8. Embauché(e) → Onboarding.

## 6. Contraintes techniques

- Excel 2019 / 365, **aucune macro**. Fonctions robustes uniquement : INDEX/MATCH, COUNTIFS, SUMIFS, AVERAGEIFS, MAXIFS, IFERROR, HYPERLINK, OFFSET (plage nommée). **Pas** de FILTER/SORT/UNIQUE/XLOOKUP (compatibilité et fichiers partagés).
- Tout calcul par formule (aucun résultat collé en dur) ; zéro erreur #N/A, #DIV/0!, #VALUE! même classeur vide.
- Plages nommées pour toutes les listes, seuils et filtres ; validations de données sur chaque champ à choix ; dates jj/mm/aaaa ; montants « # ##0 € ».
- Formats conditionnels : verdicts (vert / orange / rouge), notes 1→5 (rouge → vert), barres de données sur les scores, alertes en rouge, statut de poste, avancement d'onboarding.
- Sélection Scorecard pilotée par l'**ID** (4 premiers caractères du libellé) pour rester stable si on trie le Pipeline.

## 7. Design & ergonomie

- Police Arial ; quadrillage masqué ; fond chaud #F6F4F0 ; bandeau titre coloré par famille d'onglet (navy #1B3A5C, corail #E76F51, turquoise #2A9D8F, lavande #8E7CC3, ocre #B7791F, vert #3D8B5A).
- **Cellules à saisir : texte bleu sur fond bleu clair** ; calculs sur fond gris ; ▼ dans l'en-tête des colonnes à liste ; légende sur l'accueil.
- Couleur par poste partout : EJE lavande, AP turquoise, CAP AEPE corail.
- Lien « ⌂ Accueil » en haut de chaque onglet, volets figés sur les en-têtes, mise en page d'impression (Scorecard tenant en largeur A4 paysage, en-tête répété).

## 8. Données

Livre **2 versions** : une vierge avec **une seule ligne d'exemple** (C001, à effacer), et une version **démo** avec ~20 candidats fictifs répartis sur les 2 crèches, les 3 postes et toutes les étapes (dont au moins : 1 VETO, 1 évaluation partielle, 1 dossier incomplet en proposition, 2 relances, 3 embauches), pour voir les dashboards vivants.

## 9. Contrôle qualité avant livraison

- [ ] Recalcul complet : 0 erreur de formule (classeur démo ET classeur vierge vidé).
- [ ] Un score recalculé à la main correspond (ex. notes 5,4,5,5,4,4,5,4,4,5,4,4 pour un EJE = **88,5**).
- [ ] Changer le filtre du Dashboard modifie KPI et graphiques ; changer le candidat de la Scorecard modifie questions, notes, radar et verdict.
- [ ] Un 1 en Q3 déclenche le VETO ; une attestation d'honorabilité « Demandée » à l'étape 7 déclenche l'alerte de blocage.
- [ ] Le contrôle « ≥ 40 % de catégorie 1 » réagit aux embauches.
- [ ] Tous les liens de navigation fonctionnent ; aucun texte tronqué dans les en-têtes.

Livre le fichier, puis en 5 lignes maximum : comment démarrer et ce qu'il faut personnaliser en premier.
