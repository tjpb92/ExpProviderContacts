# ExpProviderContacts
Programme Java permettant d?exporter d?une base de données Mongo DB locale vers un fichier Excel des intervenants (ProviderContacts)

## Utilisation:
```
java ExpProviderContacts [-mgodb prod|pre-prod] [-u unum | -clientCompany clientCompanyUUID] [-p chemin vers fichier] [-o fichier] [-d] [-t] 
```
où :
* ```-mgodb prod|pre-prod``` est la référence à la base de données MongoDb, par défaut désigne la base de données de pre-production. Voir fichier *ExpProviderContacts.prop* (optionnel).
* ```-p chemin vers fichier``` est le chemin vers le fichier Excel. Amorcé à vide par défaut (paramètre optionnel).
* ```-o fichier``` est le nom du fichier Excel qui recevra les sociétés. Amorcé à *providerContacts.xlsx* par défaut (paramètre optionnel).
* ```-u unum``` est l'identifiant interne du client concerné. Non définit par défaut (paramètre optionnel).
* ```-clientCompany uuid``` est l'identifiant universel unique du client concerné. Non définit par défaut (paramètre optionnel).
* ```-d``` le programme s'exécute en mode débug, il est beaucoup plus verbeux. Désactivé par défaut (paramètre optionnel).
* ```-t``` le programme s'exécute en mode test, les transactions en base de données ne sont pas faites. Désactivé par défaut (paramètre optionnel).

## Pré-requis :
- Java 6 ou supérieur.
- JDBC Informix
- Driver MongoDB
- [xmlbeans-2.6.0.jar](https://xmlbeans.apache.org/)
- [commons-collections4-4.1.jar](https://commons.apache.org/proper/commons-collections/download_collections.cgi)
- [junit-4.12.jar] (https://github.com/junit-team/junit4/releases/tag/r4.12)
- [hamcrest-2.1.jar] (https://search.maven.org/search?q=g:org.hamcrest)

## Fichier des paramètres : 

Ce fichier permet de spécifier les paramètres d'accès aux différentes bases de données.

A adapter selon les implémentations locales.

Ce fichier est nommé : *ExpProviderContacts.prop*.

Le fichier *ExpProviderContacts_Example.prop* est fourni à titre d'exemple.

## Références:

- [API Java Exel POI](http://poi.apache.org/download.html)
- [Tuto Java POI Excel](http://thierry-leriche-dessirier.developpez.com/tutoriels/java/charger-modifier-donnees-excel-2010-5-minutes/)
- [Tuto Java POI Excel](http://jmdoudoux.developpez.com/cours/developpons/java/chap-generation-documents.php)

