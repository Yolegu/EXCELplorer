# Objectifs d'EXCELplorer

La macro EXCELplorer doit, à partir d'un fichier source au format .des, fournir un graphique vectoriel généré grâce à Excel et le copier dans un document Word. Le graphique ainsi généré doit être : 
* redimensionnable
* modifiable (chnager les couleurs de courbes, des points...)
* exportable dans un document Powerpoint
* exportable dans un document PDF sans perte de qualité (le format vectoriel doit être conservé)

# Versions d'Office utilisables

La macro EXCELplorer fonctionne sous Office 2013 et 2016. A cause d'importantes modifications du langage VBA entre les versions 2010 et 2013, une version d'EXCELplorer fonctionnant avec des versions antérieures d'Office n'est pas en développement.

L'utilisation de la macro nécessite l'installation des logiciels suivants : 
* Word
* Excel
* Powerpoint

# Téléchargement d'EXCELplorer

La code VBA est intégralement disponible dans le document [EXCELplorer.vba](https://github.com/Yolegu/EXCELplorer/blob/master/EXCELplorer.vba) du repository. Pour l'utiliser, il faut coller ce code dans un module VBA. Le document [EXCELplorer.xlsm](https://github.com/Yolegu/EXCELplorer/blob/master/EXCELplorer_v0.2.1.xlsm) contient le code d'EXCELplorer ainsi qu'une interface graphique pour l'exécuter permettant une utilisation simplifiée de la macro.

# Utilisation d'EXCELplorer

Lorsque la macro est exécutée dans EXCEL, une première fenêtre de sélection de fichier s'ouvre. Il s'agit là de sélectionner le fichier Word dans lequel les graphiques vont être crées. Si ce document n'existe pas, il est possible de le créer à cette étape. Si le document existe déjà, il suffit de la sélectionner et d'appuyer sur "Entrée".

S'ouvre alors une seconde fenêtre. Il faut cette fois sélectionner les fichiers *.des contenant les données sources des graphiques. Plusieurs documents peuvent être sélectionnées en maintenant la touche "Ctrl" enfoncée au moment de la sélection. Appuyer sur "Entrée" une fois la sélection effectuée.

A partir de ce moment, la macro génère les graphiques de la manière suivante : 
* les données sources sont lues et un graphique est généré dans Excel. Les éléments du graphique (taille du texte, position des flèches, couleurs des traits) sont définis à cette étape. Seule la mise en indice ou exposant des caractères n'est pas faite dans Excel car cela n'est pas possible.
* les graphiques sont copiés un à un dans un nouveau document Powerpoint. Le document Powerpoint créer contient une unique slide contenant un unique graphique. Chaque document Powerpoint est momentanément sauvegardé dans le répertoire contenant le fichier Word sélectionné au préalable. C'est à cette étape du processus que les indices et exposants sont traités.
* chaque graphique Powerpoint créé est ensuite copié dans le document Word spécifié à la première étape de l'exécution de la macro. Le graphique Excel est inséré dans un OLE "Diapositive Microsoft Powerpoint". Le dernier graphique inséré est placé en tête du document.
* les graphiques générés avec Excel ainsi que les fichiers Powerpoint temporaires sont finalement supprimés. Au final, seul le document Word sélectionné par l'utilisateur à été modifié.

# Choix du séparateur décimal

Le séparateur décimal des graphiques est celui défini dans les options régionales de l'ordinateur.

# Indices et exposants

Les indices et exposants peuvent être utilisés dans les légendes :
* indice : "texte normal^{mon texte en indice}"
* exposant : "texte normal^{mon texte en exposant}"

# Codes pour les marqueurs

![alt text](https://github.com/Yolegu/EXCELplorer/blob/master/img/couleurs_marqueurs.png?raw=true)
   
# Caractères grecs

| Majuscule | Code | Minuscule | Code |
| ----------| -----| ----------| -----|
| Α	| \Alpha	| α	| \alpha| 
| Β	| \Beta		| β	| \beta| 
| Γ	| \Gamma		| γ	| \gamma| 
| Δ	| \Delta	| 	δ	| \delta| 
| Ε	| \Epsilon| 		ε| 	\epsilon| 
| Ζ	| \Zeta		| ζ	| \zeta| 
| Η	| \Eta	| 	η	| \eta| 
| Θ	| \Theta		| θ	| \theta| 
| Ι	| \Iota		| ι	| \iota| 
| Κ	| \Kappa		| κ	| \kappa| 
| Λ	| \Lambda		| λ	| \lambda
| Μ	| \Mu		| μ	| \mu| 
| Ν	| \Nu		| ν	| \nu| 
| Ξ	| \Xi		| ξ	| \xi| 
| Ο	| \Omicron		| ο	| \omicron| 
| Π	| \Pi		| π	| \pi| 
| Ρ	| \Rho		| ρ	| \rho| 
| Σ	| \Sigma		| σ	| \sigma| 
| Τ	| \Tau		| τ	| \tau| 
| Υ	| \Upsilon		| υ	| \upsilon| 
| Φ	| \Phi		| φ	| \phi| 
| Χ	| \Chi		| χ	| \chi| 
| Ψ	| \Psi		| ψ	| \psi| 
| Ω	| \Omega		| ω	| \omega| 

# Redimensionnement des graphiques dans Word

# Édition des graphiques dans Word
