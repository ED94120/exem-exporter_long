# **Extraction capteurs EXEM (séries longues)**

## **Objectif**

Ce script a une finalité exclusivement **informative, éducative et de recherche**.
Il permet d’analyser l’évolution temporelle des niveaux d’exposition mesurés par les capteurs EXEM affichés sur le site EXEM de l’Observatoire des ondes (https://www.observatoiredesondes.com/fr/).
L’objectif est d’étudier le comportement des niveaux d’exposition en fonction du temps et de l’environnement du capteur.

Le script est conçu pour respecter les modalités d’utilisation du site ANFR :
- aucune requête automatisée vers les serveurs n’est effectuée,
- aucune base de données ni API n’est interrogée,
- seules les informations déjà visibles à l’écran, en particulier celles du Graphique Temps - mesures, sont lues localement dans le navigateur.

## **Conditions de fonctionnement**

Le script fonctionne quelle que soit la période de temps affichée sur le Graphique :
- que **les points de mesure sur le Graphique soient visibles ou non **,
- la période affichée soit courte ou longue,
- Il est tout à fait possible d'extraire les données à partir de la date de mise en service du Capteur.

## **Principe technique**

Le programme :
- lit les pixels qui matérialisent la courbe Temps - mesures sur le Graphique,
- convertit ces pixels en valeurs de date (jour, heures, minutes) et de niveau d'Exposition en V/m,
- calcule les statistiques (min, moyenne, max, etc.),
- génère un fichier CSV structuré.

Le script ne lit aucune donnée cachée et n’accède pas aux serveurs du site.
Il se contente d’automatiser la lecture des informations déjà affichées à l’écran.

## Glossaire des indicateurs statistiques calculés

La définition des champs du CSV (META / STAT_HORAIRE / STAT_ANNUELLE / STAT_HOURP95 / AUDIT, etc.) est disponible ici :

https://sites.google.com/view/cemethconseil/calculs-en-ligne/extraction-capteurs-exem?authuser=0#h.3aa2m9sn49gx
