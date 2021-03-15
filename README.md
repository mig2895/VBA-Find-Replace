# VBA-Find-Replace

## Fonctionnement
Une fois le logiciel lancé par le bouton sur la première page, le logiciel va d’abord sélectionner les données sur les deux colonnes, nettoyer ces données, puis les lier dans un tableau. Ensuite,  la fonction GetFolder() sera lancée et permet d’ouvrir une fenêtre permettant à l’utilisateur de choisir le dossier qu’il veut traiter. Cette fonction va donc stocker le chemin depuis la base de l’ordinateur jusqu’au dossier sélectionné. 

Avec ce nouveau chemin le logiciel va appeler la fonction LoopAllSubFolders() qui par une suite de boucles va parcourir tous les fichiers et sous dossiers. Une fois qu’il aura trouvé un fichier il va vérifier si le fichier est bien un document .docx, si c’est le cas alors il lancera d’abord la fonction DocSearch() puis la Fonction CloseWordDocuments().

La fonction Docsearch() est divisée en 3 étapes. Dans un premier temps la fonction va ouvrir le document sélectionné avec Word. Une fois le document ouvert il va chercher puis remplacer les termes appropriés, avant de sauvegarder le document édité. La fonction CloseWordDocuments() sera ensuite lancé afin de fermer toute instance du programme Word. Cette étape est nécessaire pour ne pas ouvrir une multitude de fois le même logiciel. Une fois cette fonction finie, le logiciel passera au prochain fichier qu’il a trouvé , jusqu’à ce qu’il n’en trouve plus.

Finalement, une fois que tous les documents auront été traités, le logiciel ouvrira une fenêtre indiquant à l’utilisateur qu’il a fini. 

## Limites et améliorations possibles
S’il devait y avoir une deuxième version du logiciel la première chose à faire serait de faire une recherche intelligente des termes à rechercher. En effet, actuellement le logiciel cherche exactement le terme demandé, c’est-à-dire que si un utilisateur commet une erreur de frappe le logiciel ne reconnaitra pas le terme.  Il faudrait donc intégrer une recherche intelligente (aussi connu comme « fuzzy search ») capable de comprendre que si un terme ressemble à 90% au terme recherché,  il faut les considérer comme valable.

En termes de UX, nous pensons aussi qu'une interface graphique serait l'une des fonctionnalités les plus pratiques pour l'utilisateur. En effet, celle-ci permettrrait une plus grande clareté dans l'utilisation du logiciel ainsi que potentiellement une base sur laquelle d'autres fonctionnalités pouraient être ajoutés par la suite. Par exemple, chocher une case pour choisir une recherche simple ou intelligente. 

Une autre amélioration possible est la compatibilité avec différents types de documents. En analysant les données recueillis pour ce projet il se trouve que la majorité des documents étaient de type .docx, c’est pourquoi nous nous sommes focalisés sur la solution qui convienne le mieux à ce format. Néanmoins, pour pouvoir reprendre cet outil pour d’autres projets il faudra assurer une compatibilité avec plus de formats. En effet, nous pensons notamment aux alternatives Word tels que le format OpenDocumentFormat (ODF) mais également les formats plus courants tel que le PDF. 

Finalement, nous pourrions également envisager de réunir les deux logiciels crées afin de créer un produit tout en un. C’est-à-dire qu’une fois lancé, le logiciel pourra non seulement pseudonimiser les documents mais également enlever les balises de créateur/éditeur, exporter les documents en PDF, puis répertorier tous ces fichiers sous un nouveau document avec des liens. 


## Répertoire Documents

Une fois les documents traités, il fallait un moyen de les répertorier facilement. Il serait possible de créer un tableau regroupant tous les fichiers manuellement mais cela prendrait beaucoup de temps. J’ai donc décidé de reprendre ce que j’ai appris pour le premier logiciel et créer un deuxième logiciel qui permette de créer une liste des fichiers à partir d’un dossier et sous dossiers. 

En développant le premier logiciel avec une base de plusieurs fonctions séparées, j’ai pu reprendre certaines de ces fonctions pour développer ce deuxième logiciel. En effet les fonctionnalités permettant d’ouvrir une fenêtre pour que l’utilisateur indique le dossier à traiter ainsi que la boucle permettant de traiter tous les fichiers d’un dossier, ont pu être reprises pour ce projet. La seule grande différence est la fonction HyperlinkFileListing() qui permet de parcourir les différents fichiers et les répertorier sur la première colonne de la feuille Excel. Il crée également un lien menant directement vers le fichier. 

Ne voulant pas que ce logiciel soit uniquement utilisable pour ce projet j’ai pris la décision d’énumérer les fichiers sur la première colonne et non de les trier immédiatement. En effet, cela permet de reprendre ce logiciel afin de répertorier toute sorte de fichiers pour différents projets et fichiers qui n’auraient pas le même format de nomenclature. Par conséquent, cette dernière étape de répartition des documents dans un tableau à été faite à la main. 
