# ExpProviderContacts
Programme Java permettant d?exporter d?une base de donn�es Mongo DB locale vers un fichier Excel des intervenants (ProviderContacts)

## Utilisation:
```
java ExpProviderContacts [-mgodb prod|pre-prod] [-u unum | -clientCompany clientCompanyUUID] [-p chemin vers fichier] [-o fichier] [-d] [-t] 
```
o� :
* ```-mgodb prod|pre-prod``` est la r�f�rence � la base de donn�es MongoDb, par d�faut d�signe la base de donn�es de pre-production. Voir fichier *ExpProviderContacts.prop* (optionnel).
* ```-p chemin vers fichier``` est le chemin vers le fichier Excel. Amorc� � vide par d�faut (param�tre optionnel).
* ```-o fichier``` est le nom du fichier Excel qui recevra les soci�t�s. Amorc� � *providerContacts.xlsx* par d�faut (param�tre optionnel).
* ```-u unum``` est l'identifiant interne du client concern�. Non d�finit par d�faut (param�tre optionnel).
* ```-clientCompany uuid``` est l'identifiant universel unique du client concern�. Non d�finit par d�faut (param�tre optionnel).
* ```-d``` le programme s'ex�cute en mode d�bug, il est beaucoup plus verbeux. D�sactiv� par d�faut (param�tre optionnel).
* ```-t``` le programme s'ex�cute en mode test, les transactions en base de donn�es ne sont pas faites. D�sactiv� par d�faut (param�tre optionnel).

## Pr�-requis :
- Java 6 ou sup�rieur.
- JDBC Informix
- Driver MongoDB
- [xmlbeans-2.6.0.jar](https://xmlbeans.apache.org/)
- [commons-collections4-4.1.jar](https://commons.apache.org/proper/commons-collections/download_collections.cgi)
- [junit-4.12.jar] (https://github.com/junit-team/junit4/releases/tag/r4.12)
- [hamcrest-2.1.jar] (https://search.maven.org/search?q=g:org.hamcrest)

## Fichier des param�tres : 

Ce fichier permet de sp�cifier les param�tres d'acc�s aux diff�rentes bases de donn�es.

A adapter selon les impl�mentations locales.

Ce fichier est nomm� : *ExpProviderContacts.prop*.

Le fichier *ExpProviderContacts_Example.prop* est fourni � titre d'exemple.

## R�f�rences:

- [API Java Exel POI](http://poi.apache.org/download.html)
- [Tuto Java POI Excel](http://thierry-leriche-dessirier.developpez.com/tutoriels/java/charger-modifier-donnees-excel-2010-5-minutes/)
- [Tuto Java POI Excel](http://jmdoudoux.developpez.com/cours/developpons/java/chap-generation-documents.php)

