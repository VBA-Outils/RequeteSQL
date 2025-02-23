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
<h2>Créer les domaines de valeurs pour chaque attribut.</h2>
<p>Dans l'onglet DV, pour chaque attribut, créer une ligne avec la valeur "Non renseigné" ou "Non renseignée" en fonction du genre. Puis insérer autant de lignes que de valeurs possibles.</p>
<p>K=La colonne B contient le libellé de la valeur de l'attribut, la colonne C contient la valeur technique stockée en BDD. Si l'attribut ne sert que pour restiteur les données de l'entité, alors renseigner "Oui" et "Non" en colonne B, et respectivement "RESTITUE" et "NON RESTITUE" en colonne C. Si l'attribut ne contient pas de listes de valeurs, par exemple pour une quantité saisie, alors seule la première ligne "Non renseigné(e)" est créée.</p>
<p>Lorsque tous les domaines de valeurs ont été créés, cliquer sur le bouton "Réactualiser les noms Excel" qui va créer les noms utilisés dans les listes déroulantes, et mettre à jour les colonnes E (création de la liste déroulante) et F (recherche la valeur technique).</p>
<p>La colonne H contient les contrôles qui doivent être réalisés sur les attributs, notamment des dépendances entre attributs. En l'absence de contrôle particulier, la formule =SI(Fxx="";"Absence de données "&Cxx&". Requête impossible à créer.";"") peut être utilisée en remplaçant xx par le numéro de la ligne</p>
<p>La colonne G vérifie si une erreur est présente en colonne H et incrémente le compteur d'erreur. La formule =SI(Fxx="";1;0 + SI(Hxx<>"";1;0)) doit être utilisée en remplaçant xx par le numéro de ligne</p>
<h1>Corriger la macro de génération de la requête</h1>
