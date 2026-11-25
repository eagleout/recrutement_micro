/*
  Configuration de l'application hébergée (GitHub Pages, serveur, fichier local).

  - Laissé vide : les données sont enregistrées dans le navigateur (un seul appareil).
    Pensez à faire une « Sauvegarde complète (JSON) » régulière depuis Réglages.
  - Rempli avec un projet Supabase : données partagées entre appareils et membres de l'équipe,
    connexion par lien e-mail, accès limité aux adresses listées dans la table « membres ».
    Voir supabase/schema.sql et le README.

  La clé « anon » est publique par conception : la protection des données repose sur les
  règles d'accès (RLS) créées par schema.sql. Ne mettez JAMAIS ici la clé « service_role ».

  Ce fichier n'est pas utilisé quand l'application tourne dans Claude.
*/
window.RECRUT_CONFIG = {
  supabaseUrl: "",      // ex. "https://abcdefghijkl.supabase.co"
  supabaseAnonKey: ""   // ex. "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9..."
};
