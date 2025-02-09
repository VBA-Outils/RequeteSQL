# RequeteSQL
<h1>Générateur de requêtes SQL</h1>
<p>Le fichier Excel "Requêtes.xlsm" est une base afin de créer un générateur de requête SQL à partir d'informations saisies dans une feuille Excel.</p>
<p>La feuille de saisie est découpée fonctionnellement par entités (tables) et attributs (colonnes) afin de pouvoir soit restituer des informations de l'entité soit filtrer les résultats selon les critères sélectionnés.</p>
<p>La feuille de saisie est protégée afin que les utilisateurs ne puissent modifier que les informations nécessaires à la création des requêtes.</p>
<h1>Remplir le fichier Excel</h1>
<h2>Activer le mode Développeur</h2>
<p>Afin de pouvoir modifier le fichier, le mode "Développeur" doit être activé. Aller dans l'onglet "Paramètres", puis modifier le paramètre "Mode développeur" en "Vrai".</p>
<p>Cette action ôte la protection de la feuille de saisie et affiche la colonne "G" qui est masquée en mode "Utilisateur".</p>
<h2>Découper par entité</h2>
<p>La première étape est de renseigner toutes les entités qui pourront être restituées (SELECT) ou utilisées comme critères de recherche (clause WHERE).</p>
<p>Dans l'exemple fourni, les entités "Bouteille" et "Composition" ont été créées.</p>
<p>L'entité Bouteille est constituée des données suivantes : Restituer Bouteille, Marque, Ville. Pour chaque entité, il est possible de prévoir une donnée "Restituer Entité" qui permettra de restituer la liste des attributs dans le résultat de la requête si la valeur choisie est égale à "Oui". Les autres attributs seront utilisés comme crtières de sélection dans la clause WHERE de la requête.</p>

