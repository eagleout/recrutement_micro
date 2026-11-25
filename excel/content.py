# -*- coding: utf-8 -*-
"""Contenu métier : fiches de poste, banque de questions, scripts, modèles."""

BLOCS = [
    # code, libellé long, libellé court, éliminatoire, ce qu'on évalue
    ("B1", "Sécurité & santé de l'enfant", "Sécurité", "Oui",
     "Vigilance, protocoles d'urgence, hygiène, PAI, gestes qui sauvent"),
    ("B2", "Bientraitance & posture professionnelle", "Bientraitance", "Oui",
     "Respect de l'enfant, prévention des violences éducatives ordinaires, régulation émotionnelle"),
    ("B3", "Éveil & développement de l'enfant", "Éveil & dév.", "Non",
     "Connaissance du développement 0-3 ans, activités, autonomie, observation"),
    ("B4", "Relation aux familles", "Familles", "Non",
     "Accueil, transmissions, gestion des désaccords, confidentialité"),
    ("B5", "Travail en équipe & communication", "Équipe", "Non",
     "Coopération, gestion de conflit, adaptabilité, (EJE : animation d'équipe)"),
    ("B6", "Expérience, motivation & adéquation", "Exp. & motiv.", "Non",
     "Parcours, projet, disponibilités, adéquation micro-crèche"),
]

# Pondérations par poste (somme = 100 %)
POIDS = {
    "EJE":      [0.15, 0.20, 0.25, 0.15, 0.15, 0.10],
    "AP":       [0.25, 0.20, 0.15, 0.15, 0.15, 0.10],
    "CAP AEPE": [0.20, 0.25, 0.15, 0.10, 0.15, 0.15],
}

POSTES = [
    # code, intitulé, catégorie, nb par crèche, diplôme requis
    ("EJE", "Éducateur·rice de jeunes enfants", 1, 1, "DE Éducateur de jeunes enfants"),
    ("AP", "Auxiliaire de puériculture", 1, 1, "DE Auxiliaire de puériculture"),
    ("CAP AEPE", "Accompagnant·e éducatif·ve petite enfance", 2, 2, "CAP AEPE (ou qualification cat. 2)"),
]

ETAPES = [
    # libellé, SLA jours, responsable, livrable, libellé court
    ("1. Candidature reçue", 2, "Gestionnaire", "Saisie au Pipeline + accusé de réception", "Candidatures"),
    ("2. Présélection CV", 2, "Gestionnaire / Direction", "Diplôme conforme ? Expérience ? → go / no go", "Présélection"),
    ("3. Entretien téléphonique", 3, "Gestionnaire", "Script 15 min (onglet Scripts) + note /5", "Entretien tél."),
    ("4. Entretien + mises en situation", 5, "Direction + EJE / gestionnaire", "Scorecard 12 questions notée /100", "Entretien"),
    ("5. Rencontre équipe / visite", 5, "Direction", "Visite + échange équipe (sans prise en charge d'enfants)", "Rencontre équipe"),
    ("6. Références & conformité", 7, "Gestionnaire", "2 références + diplôme + attestation d'honorabilité + aptitude", "Conformité"),
    ("7. Proposition envoyée", 5, "Gestionnaire", "Promesse d'embauche / contrat", "Proposition"),
    ("8. Embauché(e)", 0, "Gestionnaire", "DPAE + lancement de l'onboarding", "Embauche"),
]

STATUTS = ["En cours", "Embauché(e)", "Vivier", "Refusé(e)", "Désisté(e)"]
SOURCES = ["Indeed", "France Travail", "Hellowork", "LinkedIn", "Cooptation",
           "École / organisme de formation", "Candidature spontanée",
           "Facebook / groupes locaux", "Agence d'intérim", "Réseau interne", "Autre"]
DIPLOMES = ["DE Éducateur de jeunes enfants", "DE Auxiliaire de puériculture", "CAP AEPE",
            "Titre pro Intervenant éducatif petite enfance", "Bac pro ASSP",
            "DE Infirmier / Puériculteur", "Autre / en cours"]
CONTRATS = ["CDI", "CDD", "Apprentissage", "Intérim"]
OUINON = ["Oui", "Non", "En attente"]
HONOR = ["Reçue", "Demandée", "Non demandée"]
VACCINS = ["OK", "En attente"]
REFS = ["OK", "Réserve", "Négatif", "Non contactées"]
NOTES = [1, 2, 3, 4, 5]
ONB = ["✔ Fait", "En cours", "À faire", "N/A"]

ECHELLE = [
    (1, "Insuffisant / alerte", "Réponse dangereuse, inadaptée ou absente. Une note de 1 en Sécurité ou Bientraitance déclenche un VETO."),
    (2, "Fragile", "Réponse partielle, notions floues : nécessiterait un accompagnement important."),
    (3, "Conforme", "Réponse correcte : l'attendu de base est maîtrisé."),
    (4, "Solide", "Réponse structurée, exemples concrets vécus, recul sur sa pratique."),
    (5, "Excellent", "Réponse experte : anticipe, prévient, sait transmettre aux autres."),
]

PREREQUIS = ("• Copie du diplôme (vérifiée auprès de l'organisme si doute)\n"
             "• Attestation d'honorabilité (honorabilite.social.gouv.fr) — obligatoire depuis le 01/10/2025 "
             "à l'embauche puis périodiquement ; l'employeur la récupère, la vérifie et la conserve\n"
             "• Vaccinations / aptitude : à valider avec la médecine du travail\n"
             "• 2 références professionnelles contactées\n"
             "• Pièce d'identité, justificatifs administratifs, DPAE avant le 1er jour")

FICHES = {
    "EJE": {
        "titre": "ÉDUCATEUR·RICE DE JEUNES ENFANTS (EJE)",
        "Catégorie réglementaire": "Catégorie 1 — diplômé·e d'État : compte dans les 40 % minimum de l'effectif (décret n° 2025-304).",
        "Diplôme requis": "Diplôme d'État d'éducateur de jeunes enfants (DEEJE) — obligatoire.",
        "Rattachement": "Direction de la micro-crèche / gestionnaire. Référent·e éducatif·ve de l'équipe (1 AP + 2 CAP AEPE).",
        "Raison d'être du poste": "Garant·e de la qualité éducative : faire vivre le projet pédagogique, accompagner le développement de chaque enfant, soutenir l'équipe et les familles.",
        "Missions principales": [
            "Concevoir, animer et évaluer le projet pédagogique avec la direction",
            "Accueillir enfants et familles, conduire les familiarisations et les entretiens familles",
            "Proposer des activités d'éveil adaptées à l'âge et au rythme de chacun (jeu libre, motricité libre, langage, sensoriel, éveil culturel)",
            "Observer, rédiger des écrits professionnels, repérer les difficultés de développement et alerter",
            "Garantir la sécurité physique et affective ; faire appliquer les protocoles (hygiène, urgence, PAI)",
            "Animer la réflexion d'équipe : réunions, transmissions, analyse des pratiques, tutorat des stagiaires et apprentis",
            "Participer aux soins quotidiens (repas, change, sieste) : en micro-crèche, l'EJE est aussi sur le terrain",
            "Développer les partenariats (PMI, RPE, médiathèque…) et l'inclusion des enfants à besoins particuliers",
        ],
        "Compétences techniques": [
            "Développement psychomoteur, affectif et cognitif 0-3 ans",
            "Pédagogies de référence (Pikler-Lóczy, Montessori, motricité libre) et Charte nationale pour l'accueil du jeune enfant",
            "Observation, écrits professionnels, conduite de projet",
            "Animation d'équipe et feedback constructif",
            "Prévention des violences éducatives ordinaires (VEO)",
        ],
        "Savoir-être": "Leadership bienveillant · posture réflexive · créativité · stabilité émotionnelle · discrétion · sens de l'organisation",
        "Expérience souhaitée": "1 à 3 ans en EAJE (débutant·e accepté·e si stages solides). Une expérience de tutorat ou d'encadrement est un plus.",
        "Conditions de travail": "CDI 35 h · horaires tournants sur l'amplitude d'ouverture (ex. 7h30-19h) · travail au sol et port d'enfants · réunions d'équipe hors temps d'accueil.",
        "Rémunération indicative": "≈ 2 200 – 2 800 € brut/mois (repère Île-de-France, secteur privé) — à caler sur votre grille ou convention collective.",
        "Évolution": "Coordination pédagogique, direction de micro-crèche (conditions de qualification et d'expérience à vérifier avec la PMI).",
        "Annonce prête à publier": (
            "[Nom de la crèche] recrute son·sa Éducateur·rice de jeunes enfants (H/F) — CDI 35 h — [Ville]\n"
            "Rejoignez une micro-crèche de 12 berceaux à taille humaine. Vous serez la référence éducative d'une équipe de 4 : "
            "projet pédagogique, éveil, familiarisations, lien avec les familles et accompagnement de l'équipe.\n"
            "Profil : DEEJE obligatoire, sens du collectif, créativité, posture bienveillante. Jeunes diplômé·e·s bienvenu·e·s.\n"
            "Nous offrons : une vraie place dans les choix pédagogiques, une équipe soudée, la formation continue, [avantages].\n"
            "Rémunération : [xxx] € brut/mois selon expérience · Prise de poste : [date].\n"
            "Attestation d'honorabilité exigée à l'embauche (obligation légale)."),
    },
    "AP": {
        "titre": "AUXILIAIRE DE PUÉRICULTURE (AP)",
        "Catégorie réglementaire": "Catégorie 1 — diplômé·e d'État : compte dans les 40 % minimum et sécurise la présence obligatoire d'au moins un professionnel cat. 1.",
        "Diplôme requis": "Diplôme d'État d'auxiliaire de puériculture (DEAP) — obligatoire.",
        "Rattachement": "Direction de la micro-crèche ; travail en binôme avec l'EJE. Référent·e santé / soins de l'équipe au quotidien.",
        "Raison d'être du poste": "Assurer le bien-être, la santé et la sécurité de chaque enfant par des soins de qualité et une vigilance de tous les instants.",
        "Missions principales": [
            "Accueillir enfants et parents, réaliser des transmissions orales et écrites fiables",
            "Assurer les soins d'hygiène et de confort (change, toilette, habillage) en respectant le rythme et l'intimité de l'enfant",
            "Organiser les repas (biberons, diversification, régimes, allergies) et l'endormissement",
            "Surveiller l'état de santé : repérer les signes d'alerte, appliquer les protocoles de la structure et les PAI",
            "Respecter et faire respecter l'hygiène (biberonnerie, matériel, HACCP) et la sécurité des espaces",
            "Proposer des jeux et activités d'éveil, favoriser l'autonomie et la motricité libre",
            "Participer au projet pédagogique, aux réunions et à l'accueil des stagiaires",
        ],
        "Compétences techniques": [
            "Soins de puériculture et hygiène",
            "Repérage des signes cliniques (fièvre, déshydratation, détresse respiratoire, réaction allergique)",
            "Gestes d'urgence (PSC1 / SST appréciés), protocoles et PAI",
            "Diététique du jeune enfant",
            "Développement de l'enfant 0-3 ans",
        ],
        "Savoir-être": "Vigilance · douceur · rigueur · calme en situation d'urgence · sens de l'équipe · discrétion",
        "Expérience souhaitée": "Débutant·e accepté·e ; 1 à 2 ans en crèche, maternité ou pédiatrie appréciés.",
        "Conditions de travail": "CDI 35 h · horaires tournants (ouverture / fermeture) · port d'enfants, position basse · tenue professionnelle.",
        "Rémunération indicative": "≈ 1 850 – 2 500 € brut/mois (repère offres Île-de-France 2025-2026, secteur privé).",
        "Évolution": "Référent·e santé, formation EJE ou puériculture, fonctions de coordination.",
        "Annonce prête à publier": (
            "[Nom de la crèche] recrute un·e Auxiliaire de puériculture diplômé·e (H/F) — CDI 35 h — [Ville]\n"
            "Dans une micro-crèche de 12 enfants, vous prenez soin des tout-petits au quotidien : soins, repas, sommeil, éveil, "
            "vigilance santé et lien avec les familles, en binôme avec notre EJE.\n"
            "Profil : DEAP obligatoire, douceur, rigueur et esprit d'équipe. Débutant·e accepté·e.\n"
            "Nous offrons : petite équipe, temps d'échange sur les pratiques, formation (gestes d'urgence, VEO), [avantages].\n"
            "Rémunération : [xxx] € brut/mois selon expérience · Prise de poste : [date].\n"
            "Attestation d'honorabilité exigée à l'embauche (obligation légale)."),
    },
    "CAP AEPE": {
        "titre": "ACCOMPAGNANT·E ÉDUCATIF·VE PETITE ENFANCE (CAP AEPE)",
        "Catégorie réglementaire": "Catégorie 2 — qualification ou expérience : peut représenter au plus 60 % de l'effectif (seuil de 40 % de cat. 1 à respecter).",
        "Diplôme requis": "CAP Accompagnant éducatif petite enfance (ou qualification équivalente reconnue en cat. 2).",
        "Rattachement": "Direction de la micro-crèche ; travail sous la coordination éducative de l'EJE et en lien étroit avec l'AP.",
        "Raison d'être du poste": "Accompagner chaque enfant au quotidien dans un cadre sécurisant, chaleureux et propre, en participant à tous les temps de vie.",
        "Missions principales": [
            "Accueillir l'enfant et sa famille, assurer les transmissions",
            "Réaliser les soins du quotidien : repas, change, sieste, hygiène",
            "Animer des jeux et activités d'éveil, accompagner l'enfant vers l'autonomie",
            "Veiller à la sécurité physique et affective des enfants",
            "Assurer l'entretien et l'hygiène des locaux, jouets, linge et le réchauffage des repas (HACCP)",
            "Participer aux réunions d'équipe et à la mise en œuvre du projet pédagogique",
        ],
        "Compétences techniques": [
            "Besoins fondamentaux et rythmes de l'enfant",
            "Hygiène, sécurité et règles HACCP",
            "Techniques d'animation et d'éveil",
            "Gestes de premiers secours appréciés",
        ],
        "Savoir-être": "Bienveillance · patience · ponctualité · sens de l'équipe · adaptabilité · discrétion",
        "Expérience souhaitée": "Débutant·e accepté·e (stages, garde d'enfants, assistant·e maternel·le valorisés).",
        "Conditions de travail": "CDI 35 h · horaires tournants · port d'enfants, position basse · polyvalence entretien / soins / éveil.",
        "Rémunération indicative": "Base SMIC à ≈ 2 050 € brut/mois selon expérience (repère Île-de-France, secteur privé).",
        "Évolution": "Auxiliaire de puériculture (formation ou VAE), titre professionnel Intervenant éducatif petite enfance.",
        "Annonce prête à publier": (
            "[Nom de la crèche] recrute 2 Accompagnant·e·s éducatif·ve·s petite enfance (CAP AEPE) (H/F) — CDI 35 h — [Ville]\n"
            "Vous accompagnez les tout-petits dans tous les moments de la journée : accueil, repas, change, sieste, jeux et éveil, "
            "au sein d'une équipe de 4 professionnelles bienveillantes.\n"
            "Profil : CAP AEPE, patience, ponctualité, esprit d'équipe. Débutant·e accepté·e.\n"
            "Nous offrons : intégration accompagnée, formation continue, perspectives d'évolution (VAE auxiliaire de puériculture), [avantages].\n"
            "Rémunération : [xxx] € brut/mois · Prise de poste : [date].\n"
            "Attestation d'honorabilité exigée à l'embauche (obligation légale)."),
    },
}

FICHE_SECTIONS = ["Catégorie réglementaire", "Diplôme requis", "Rattachement", "Raison d'être du poste",
                  "Missions principales", "Compétences techniques", "Savoir-être", "Expérience souhaitée",
                  "Conditions de travail", "Rémunération indicative", "Évolution", "Prérequis à l'embauche",
                  "Annonce prête à publier"]

# Banque de questions : (bloc 1-6, type, question, réponse attendue, signaux d'alerte, relance)
QUESTIONS = {
    "EJE": [
        (1, "Mise en situation",
         "Pendant le jeu libre, un enfant de 18 mois chute et se cogne la tête contre l'angle d'un meuble. Il pleure puis se calme. Que faites-vous, étape par étape ?",
         "Confie le groupe à une collègue · évalue (conscience, vomissements, bosse, comportement) · applique le protocole · surveille dans la durée · prévient les parents · trace l'incident · informe la direction · revoit l'aménagement (angle).",
         "Minimise (« ce n'est rien ») · laisse le groupe sans adulte · n'informe pas les parents · aucune traçabilité.",
         "Quels signes vous feraient appeler le 15 ?"),
        (1, "Expérience",
         "Comment intégrez-vous la sécurité dans l'aménagement de l'espace et dans la préparation d'une activité (ex. transvasement de semoule) ?",
         "Analyse des risques (ingestion de petits objets, allergies, hygiène) · adaptation à l'âge · nombre d'enfants par adulte · surveillance active · aménagement évolutif et motricité libre sécurisée.",
         "Aucune notion de risque d'étouffement · activité inadaptée à l'âge · adulte absent pendant l'activité.",
         "Racontez une activité que vous avez dû modifier pour raison de sécurité."),
        (2, "Mise en situation",
         "Vous entendez une collègue dire sèchement à un enfant : « Si tu continues, tu vas au coin et tu n'auras pas de dessert. » Que faites-vous sur le moment, puis après ?",
         "Intervient avec tact pour protéger l'enfant (prend le relais) · échange ensuite en privé · s'appuie sur la prévention des VEO et la Charte nationale · remonte à la direction si c'est répété · propose une analyse de pratiques.",
         "« Ce n'est pas mon rôle » · banalise · humilie la collègue devant les enfants ou les parents.",
         "Et si cette collègue est plus ancienne que vous dans la structure ?"),
        (2, "Expérience",
         "Racontez une situation où vous avez senti la fatigue ou l'agacement monter face à un enfant (pleurs, opposition). Comment l'avez-vous gérée ?",
         "Lucidité sur ses émotions · passage de relais · technique de régulation · retour réflexif · aucune violence verbale ou physique.",
         "Affirme ne jamais ressentir d'agacement · réponse punitive · propos dévalorisants sur l'enfant.",
         "Qu'avez-vous changé dans votre pratique ensuite ?"),
        (3, "Projet pédagogique",
         "Présentez les grands axes du projet pédagogique que vous aimeriez porter dans une micro-crèche de 12 enfants. Sur quelles références vous appuyez-vous ?",
         "Projet cohérent (jeu libre, motricité libre, langage, autonomie, inclusion, éveil culturel / nature) · références citées (Pikler-Lóczy, Montessori, Charte nationale) · déclinaison concrète · modalités d'évaluation.",
         "Catalogue d'activités dirigées sans sens · aucune référence · rigidité.",
         "Comment embarquez-vous l'équipe et les parents dans ce projet ?"),
        (3, "Mise en situation",
         "En observant un enfant de 2 ans, vous remarquez qu'il ne pointe pas du doigt, parle très peu et évite le regard. Comment procédez-vous ?",
         "Observations objectives, répétées et écrites · partage en équipe · échange avec la direction et le référent santé · entretien avec la famille avec tact, sans diagnostic · orientation (médecin, PMI) · adaptation de l'accueil.",
         "Pose un diagnostic (« il est autiste ») · en parle aux parents lors d'une transmission rapide devant d'autres · ne fait rien.",
         "Comment formuleriez-vous la première phrase aux parents ?"),
        (4, "Mise en situation",
         "Le soir, un parent en colère vous reproche devant d'autres parents que son enfant a été mordu pour la troisième fois. Comment réagissez-vous ?",
         "Accueille l'émotion · propose de s'isoler · explique la morsure (étape du développement) · présente les mesures prises (observation, aménagement) · confidentialité (ne nomme pas l'enfant mordeur) · propose un rendez-vous · informe la direction.",
         "Nomme l'enfant mordeur · se justifie de façon défensive · promet que cela n'arrivera plus.",
         "Que dites-vous aux parents de l'enfant qui mord ?"),
        (4, "Expérience",
         "Comment construisez-vous la familiarisation (adaptation) d'un bébé de 10 mois dont la maman est très anxieuse ?",
         "Progressivité · professionnel·le référent·e · rituels et objet transitionnel · écoute du parent · transmissions détaillées · souplesse du calendrier · réassurance.",
         "Adaptation standardisée et rigide · minimise l'anxiété du parent.",
         "Que faites-vous si la maman appelle 4 fois dans la journée ?"),
        (5, "Animation d'équipe",
         "Dans l'équipe (1 AP, 2 CAP AEPE), deux collègues sont en conflit et l'ambiance se dégrade. Que mettez-vous en place ?",
         "Écoute individuelle · recentrage sur l'intérêt de l'enfant · temps d'échange médiatisé · règles de fonctionnement partagées · information de la direction · suivi dans le temps.",
         "Prend parti · évitement · autoritarisme.",
         "À quel moment passez-vous le relais à la direction ?"),
        (5, "Expérience",
         "Donnez un exemple où vous avez accompagné une collègue ou une stagiaire pour faire évoluer une pratique. Comment vous y êtes-vous pris·e ?",
         "Méthode : observation, feedback constructif, démonstration, objectifs, valorisation, suivi.",
         "Pas d'exemple concret · approche descendante ou culpabilisante.",
         "Qu'est-ce qui a été le plus difficile ?"),
        (6, "Expérience",
         "Présentez votre parcours en EAJE (structures, âges accueillis) : qu'avez-vous appris de plus important, et pourquoi une micro-crèche aujourd'hui ?",
         "Parcours cohérent · recul sur ses apprentissages · compréhension des spécificités de la micro-crèche (polyvalence, petit effectif, proximité avec les familles).",
         "Vision floue · départs multiples non expliqués · recherche d'un poste « tranquille ».",
         "Pourquoi avez-vous quitté votre dernier poste ?"),
        (6, "Motivation",
         "Nous reprenons la structure : quelles seraient vos 3 priorités sur vos 90 premiers jours ? Les horaires tournants (ouverture / fermeture) vous conviennent-ils ?",
         "Priorités réalistes (connaître enfants et familles, cohésion d'équipe, projet pédagogique, protocoles) · disponibilité confirmée · projection dans la durée.",
         "Exigences incompatibles avec l'organisation · aucune projection.",
         "Où vous voyez-vous dans 2 ans ?"),
    ],
    "AP": [
        (1, "Mise en situation",
         "À 14 h, un enfant de 15 mois a 39,2 °C, il est grognon et moins réactif que d'habitude. Que faites-vous ?",
         "Applique le protocole fièvre de la structure · recherche les signes de gravité (taches, somnolence, gêne respiratoire) · traitement uniquement selon protocole et autorisation · appelle les parents · trace · appelle le 15 si signe de gravité · hydratation, isolement si besoin.",
         "Donne un médicament hors protocole ou sans autorisation · ne prévient pas les parents · ignore les signes de gravité.",
         "Que vérifiez-vous avant d'administrer un antipyrétique ?"),
        (1, "Mise en situation",
         "Un enfant ayant un PAI pour allergie à l'arachide présente, après le goûter, des plaques rouges et se met à tousser. Décrivez vos actions.",
         "Suit le PAI (trousse d'urgence, adrénaline si prévue) · appelle le 15 · reste avec l'enfant · prévient collègues, direction et parents · trace · analyse la cause (contrôle des aliments).",
         "Hésite ou attend l'arrivée des parents · ne sait pas où est le PAI · ne connaît pas la trousse d'urgence.",
         "Comment vous assurez-vous que toute l'équipe connaît les PAI ?"),
        (2, "Mise en situation",
         "Au repas, un enfant de 20 mois refuse de manger et repousse son assiette. Votre collègue insiste en lui mettant la cuillère dans la bouche. Que faites-vous ?",
         "Ne force jamais (violence éducative ordinaire) · propose sans contraindre · respecte l'appétit · intervient avec tact · échange avec la collègue après · informe les parents des quantités · remonte si c'est répété.",
         "Approuve le forçage · « il doit finir son assiette ».",
         "Et si ce sont les parents qui demandent qu'il finisse ?"),
        (2, "Expérience",
         "Comment respectez-vous l'intimité et le consentement de l'enfant lors des soins (change, toilette) ? Donnez un exemple concret.",
         "Prévient et verbalise avant de toucher · fait participer l'enfant · change debout selon l'âge · espace dédié, jamais en public · gestes doux · vraie disponibilité relationnelle.",
         "Soins « à la chaîne » · aucune verbalisation.",
         "Que faites-vous si un enfant refuse le change ?"),
        (3, "Connaissances",
         "Quels repères de développement moteur attendez-vous entre 6 et 18 mois, et comment aménagez-vous l'espace pour favoriser la motricité libre ?",
         "Repères (retournement, assis seul, déplacements, marche vers 12-18 mois) · ne met pas l'enfant dans une position qu'il n'a pas acquise · tapis, structures basses · pas de transat ni de trotteur.",
         "Assied des bébés qui ne tiennent pas seuls · utilise des appareils contraignants.",
         "Comment l'expliquez-vous à des parents qui veulent « le faire marcher » ?"),
        (3, "Mise en situation",
         "Un bébé de 8 mois arrivé récemment pleure beaucoup et dort mal à la crèche. Comment l'accompagnez-vous ?",
         "Observe ses rythmes · échange avec les parents sur ses habitudes · référence · objet transitionnel · portage et réassurance · environnement calme · patience · transmissions.",
         "Le laisse pleurer seul · horaires rigides pour tous.",
         "Au bout de combien de temps vous inquiéteriez-vous ?"),
        (4, "Mise en situation",
         "Un parent vous demande d'empêcher son enfant de 2 ans de faire la sieste pour qu'il dorme mieux le soir. Comment répondez-vous ?",
         "Écoute · explique le besoin de sommeil · propose un compromis (ne pas empêcher, temps de repos, ne pas prolonger) · co-construit · trace · relaie à la direction en cas de désaccord.",
         "Accepte d'empêcher l'enfant de dormir · refuse sèchement sans dialogue.",
         "Si le parent insiste fortement ?"),
        (4, "Expérience",
         "Comment réalisez-vous en 2 minutes des transmissions utiles et rassurantes au parent le soir ? Donnez un exemple.",
         "Factuel (repas, sieste, change, santé) + vécu de l'enfant (moment fort) · positif sans mentir · information sensible en privé · confidentialité.",
         "Transmissions vides (« ça a été ») · jugements · parle des autres enfants.",
         "Comment annoncez-vous un petit accident (chute, morsure) ?"),
        (5, "Mise en situation",
         "À l'ouverture, une collègue appelle : elle sera absente. Les enfants arrivent. Que faites-vous ?",
         "Vérifie immédiatement le taux d'encadrement · prévient la direction / le gestionnaire · réorganise · cherche un remplacement · ne dépasse jamais le ratio autorisé · priorité à la sécurité.",
         "Accepte de dépasser le taux sans alerter · panique · laisse un adulte seul avec trop d'enfants.",
         "Que dites-vous aux parents qui arrivent ?"),
        (5, "Expérience",
         "Racontez un désaccord avec une collègue ou une responsable sur une pratique. Comment s'est-il résolu ?",
         "Communication directe et respectueuse · arguments centrés sur l'enfant · accepte le compromis · en tire un apprentissage.",
         "Rumeurs · conflit non résolu · critique systématique des anciens employeurs.",
         "Qu'auriez-vous fait différemment ?"),
        (6, "Expérience",
         "Présentez votre parcours (crèche, maternité, pédiatrie…) : qu'est-ce qui vous attire dans une micro-crèche de 12 enfants ?",
         "Parcours cohérent · goût de la polyvalence et du travail en petite équipe · vraie motivation pour l'accueil collectif.",
         "Motivation uniquement pratique (proximité, horaires) · méconnaissance du collectif.",
         "Quelle a été votre plus grande fierté professionnelle ?"),
        (6, "Motivation",
         "Quelles sont vos disponibilités et contraintes (horaires tournants, trajet, date de démarrage) ? Comment voyez-vous votre évolution ?",
         "Disponibilités compatibles · trajet raisonnable · projet réaliste (référent·e santé, formation).",
         "Contraintes incompatibles · aucune projection.",
         "Quelle formation aimeriez-vous suivre dans l'année ?"),
    ],
    "CAP AEPE": [
        (1, "Mise en situation",
         "En fin de journée, une personne que vous ne connaissez pas se présente pour récupérer un enfant en disant être sa tante. Que faites-vous ?",
         "Ne confie jamais l'enfant à une personne non autorisée · vérifie la liste des personnes autorisées et la pièce d'identité · appelle les parents · informe la direction · reste courtois·e.",
         "Confie l'enfant « parce qu'elle a l'air de le connaître ».",
         "Et si l'enfant court vers elle en souriant ?"),
        (1, "Mise en situation",
         "Pendant que vous préparez le goûter, un enfant de 2 ans met une petite perle dans sa bouche et commence à s'étouffer. Que faites-vous ?",
         "Gestes d'urgence adaptés à l'âge (claques dans le dos, compressions) · alerte une collègue · appel au 15 / 112 · protocole · prévention : petits objets hors de portée.",
         "Ne connaît aucun geste · va chercher l'objet avec les doigts à l'aveugle · panique.",
         "Avez-vous déjà suivi une formation aux gestes de premiers secours ?"),
        (2, "Mise en situation",
         "Un enfant de 2 ans et demi tape les autres et vous mord pendant un jeu. Comment réagissez-vous ?",
         "Protège l'enfant tapé · pose une limite calmement et clairement · met des mots sur l'émotion · propose une alternative · aucune punition humiliante · observe la fréquence et en parle en équipe.",
         "Le met au coin longtemps · crie · « le mordre pour qu'il comprenne ».",
         "Que dites-vous à ses parents le soir ?"),
        (2, "Connaissances",
         "Qu'est-ce que la bientraitance pour vous ? Donnez un exemple concret de votre pratique et un exemple de violence éducative ordinaire à éviter.",
         "Respect des besoins, du rythme et de la dignité · exemples (verbaliser, laisser le temps) · VEO citées (cris, menaces, forcer à manger, moqueries, étiquettes).",
         "Ne connaît pas la notion · banalise certaines VEO.",
         "Que faites-vous si vous voyez une collègue crier sur un enfant ?"),
        (3, "Mise en situation",
         "Proposez une activité d'éveil pour 4 enfants de 18 à 30 mois : objectif, matériel, déroulé, rôle de l'adulte.",
         "Activité adaptée (transvasement, pâte à sel, parcours moteur, lecture) · objectif (sensoriel, moteur, langage) · sécurité · participation libre · adulte disponible et observateur.",
         "Activité dangereuse ou très dirigée · obligation de participer · production « pour les parents ».",
         "Et si un enfant ne veut pas participer ?"),
        (3, "Connaissances",
         "Comment accompagnez-vous un enfant vers l'autonomie (repas, habillage, propreté) sans le brusquer ?",
         "Respect du rythme de chacun · encouragements · matériel adapté · lien avec la famille · aucune pression sur la propreté.",
         "« Il doit être propre à 2 ans » · pression ou moqueries.",
         "Comment réagissez-vous à un « accident » de propreté ?"),
        (4, "Mise en situation",
         "Le matin, un papa pressé part sans dire au revoir à son enfant qui pleure. Que faites-vous avec l'enfant, puis avec le papa ?",
         "Accueille et rassure l'enfant · verbalise · en reparle au papa avec bienveillance (importance d'une séparation ritualisée) · sans jugement.",
         "Juge le parent · « il vaut mieux partir vite ».",
         "Comment formulez-vous votre remarque au papa ?"),
        (4, "Expérience",
         "Un parent vous pose une question sur la santé ou le comportement de son enfant et vous ne savez pas répondre. Comment faites-vous ?",
         "Reconnaît ses limites · n'improvise pas · relaie à l'EJE, à l'AP ou à la direction · propose un rendez-vous.",
         "Donne des conseils médicaux · invente une réponse.",
         "Donnez un exemple réel."),
        (5, "Mise en situation",
         "Ce jour-là, vous êtes en charge de l'entretien des locaux et du réchauffage des repas, mais le groupe a besoin de vous. Comment vous organisez-vous ?",
         "Priorité à la sécurité des enfants · communique avec l'équipe · respecte l'HACCP (températures, traçabilité) · s'organise avec méthode.",
         "Néglige l'hygiène · laisse les enfants sans surveillance.",
         "Quelles règles HACCP connaissez-vous ?"),
        (5, "Expérience",
         "Racontez un moment où vous avez dû vous adapter rapidement à un changement (planning, collègue absente). Qu'avez-vous fait ?",
         "Flexibilité · communication · esprit d'équipe · reste centré·e sur les enfants.",
         "Rigidité · plainte sans solution.",
         "Comment réagissez-vous à un changement de planning la veille ?"),
        (6, "Expérience",
         "Parlez-nous de vos expériences auprès de jeunes enfants (crèche, stage, garde, assistant·e maternel·le) : qu'avez-vous préféré, et trouvé difficile ?",
         "Expériences concrètes · recul · difficultés assumées et dépassées.",
         "Aucune expérience réelle décrite · discours générique.",
         "Qu'est-ce qu'une bonne journée pour vous ?"),
        (6, "Motivation",
         "Pourquoi ce métier et pourquoi notre micro-crèche ? Envisagez-vous d'évoluer (ex. auxiliaire de puériculture par VAE ou formation) ?",
         "Motivation authentique pour l'enfant · intérêt pour l'évolution (CAP → AP : un atout face aux nouvelles exigences de qualification).",
         "Choix par défaut · aucune projection.",
         "Seriez-vous prêt·e à suivre une formation en cours d'année ?"),
    ],
}

SCRIPT_TEL = [
    ("Pouvez-vous me confirmer votre diplôme exact et son année d'obtention ?", "Vérifier l'éligibilité au poste (cat. 1 / cat. 2)", "Oui", "Diplôme conforme au poste, copie transmissible rapidement."),
    ("Depuis combien de temps travaillez-vous auprès des 0-3 ans, et dans quels types de structures ?", "Qualifier l'expérience", "Non", "Expériences concrètes et datées (structures, âges, durées)."),
    ("Nos horaires tournent sur l'amplitude d'ouverture (ex. 7h30-19h). Est-ce compatible avec votre organisation ?", "Disponibilité", "Oui", "Oui clair, sans contrainte bloquante."),
    ("À quelle date pourriez-vous démarrer ? Avez-vous un préavis ?", "Planification", "Non", "Date compatible avec la prise de poste cible."),
    ("Combien de temps de trajet auriez-vous, et comment viendriez-vous ?", "Risque d'absentéisme / turnover", "Non", "Trajet raisonnable (idéalement < 45 min), solution fiable."),
    ("Quelles sont vos prétentions salariales (brut mensuel) ?", "Adéquation budget", "Non", "Dans la fourchette du poste (voir Fiches de poste)."),
    ("En une minute : pourquoi ce poste, et pourquoi une micro-crèche ?", "Motivation", "Non", "Motivation centrée sur l'enfant et le travail en petite équipe."),
    ("Nous demandons une attestation d'honorabilité à l'embauche (obligation légale). Est-ce que cela vous pose une difficulté ?", "Conformité", "Oui", "Accord sans hésitation."),
]

MODELES = [
    ("Convocation à l'entretien",
     "Objet : Votre candidature — poste de [poste] à [crèche]\n\n"
     "Bonjour [Prénom],\n\nMerci pour notre échange téléphonique. Je vous propose un entretien le [date] à [heure] à [adresse] "
     "(durée : 1 h environ, dont quelques mises en situation professionnelles).\n"
     "Merci d'apporter : CV, copie de votre diplôme et, si possible, les coordonnées de 2 références.\n\n"
     "Pouvez-vous me confirmer votre présence ?\n\nBien cordialement,\n[Signature]"),
    ("Demande des pièces de conformité",
     "Objet : Finalisation de votre recrutement — pièces à fournir\n\n"
     "Bonjour [Prénom],\n\nNous sommes heureux de poursuivre avec vous. Pour finaliser, merci de nous transmettre :\n"
     "• la copie de votre diplôme ;\n• votre attestation d'honorabilité, à demander sur honorabilite.social.gouv.fr (obligatoire pour exercer en crèche) ;\n"
     "• vos justificatifs administratifs (pièce d'identité, carte Vitale, RIB).\n"
     "Nous contacterons également les références que vous nous avez indiquées.\n\nBien cordialement,\n[Signature]"),
    ("Réponse négative après entretien",
     "Objet : Votre candidature — [poste]\n\n"
     "Bonjour [Prénom],\n\nMerci pour le temps que vous nous avez consacré lors de notre entretien. "
     "Après réflexion, nous avons retenu un profil plus proche de nos besoins actuels.\n"
     "[Option : Votre candidature a retenu notre attention ; avec votre accord, nous la conservons pour de futurs postes.]\n\n"
     "Nous vous souhaitons pleine réussite.\nBien cordialement,\n[Signature]"),
    ("Proposition d'embauche",
     "Objet : Proposition d'embauche — [poste] à [crèche]\n\n"
     "Bonjour [Prénom],\n\nNous avons le plaisir de vous proposer le poste de [poste] :\n"
     "• Contrat : [CDI] — [35 h] hebdomadaires\n• Rémunération : [xxx] € brut mensuel\n• Prise de poste : [date]\n"
     "• Période d'essai : [durée]\n\nCette proposition est valable jusqu'au [date] et reste soumise à la réception de votre attestation "
     "d'honorabilité et de votre diplôme.\n\nAu plaisir de vous accueillir dans l'équipe,\n[Signature]"),
]

ONBOARDING = [
    ("Avant J1", "Promesse d'embauche / contrat signé"),
    ("Avant J1", "DPAE effectuée"),
    ("Avant J1", "Attestation d'honorabilité vérifiée, archivée, date de renouvellement notée"),
    ("Avant J1", "Copie du diplôme archivée au dossier du personnel"),
    ("Avant J1", "Mutuelle / prévoyance, justificatifs administratifs"),
    ("Avant J1", "Visite d'information et de prévention (médecine du travail) planifiée"),
    ("Avant J1", "Planning des 4 premières semaines remis"),
    ("J1", "Accueil, visite des locaux, présentation de l'équipe"),
    ("J1", "Remise du projet pédagogique, du règlement de fonctionnement, du livret d'accueil"),
    ("Semaine 1", "Protocoles lus et signés (urgence, PAI, hygiène, sorties, sécurité)"),
    ("Semaine 1", "Binôme / tutorat avec un·e professionnel·le expérimenté·e"),
    ("Semaine 2", "Présentation aux familles des enfants dont il·elle est référent·e"),
    ("J+30", "Point « rapport d'étonnement » à 1 mois"),
    ("Fin d'essai", "Bilan de fin de période d'essai"),
    ("Fin d'essai", "Plan de formation (gestes d'urgence, prévention des VEO…)"),
]

REGLEMENTAIRE = [
    "Décret n° 2025-304 du 1er avril 2025 — dispositions micro-crèches applicables au 1er septembre 2026 :",
    "• au moins un professionnel diplômé d'État de catégorie 1 (auxiliaire de puériculture, EJE, infirmier, psychomotricien) dans l'équipe ;",
    "• les professionnels de catégorie 1 représentent au moins 40 % de l'effectif ; les CAP AEPE restent possibles dans le respect de ce seuil ;",
    "• le référent technique est remplacé par une fonction de direction d'au moins 0,5 ETP ; une même personne ne dirige pas plus de 2 structures (mesures transitoires prévues).",
    "Attestation d'honorabilité obligatoire depuis le 1er octobre 2025 pour les professionnels des crèches : à récupérer, vérifier et conserver par l'employeur, à l'embauche puis périodiquement.",
    "Votre schéma cible (1 EJE + 1 AP + 2 CAP AEPE) = 50 % de catégorie 1 : conforme au seuil de 40 %, à condition de pourvoir l'EJE et l'AP. Le Dashboard le contrôle en continu.",
]
