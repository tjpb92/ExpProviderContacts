package expprovidercontacts;

import utils.GetArgsException;
import bkgpi2a.CivilName;
import bkgpi2a.ItemAbstract;
import bkgpi2a.ItemAbstractWithRef;
import bkgpi2a.ItemAbstractWithRefList;
import bkgpi2a.Name;
import bkgpi2a.PoorName;
import bkgpi2a.ProviderContact;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.mongodb.BasicDBObject;
import com.mongodb.MongoClient;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoCursor;
import com.mongodb.client.MongoDatabase;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PaperSize;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bson.Document;
import utils.ApplicationProperties;
import utils.DBServer;
import utils.DBServerException;
import utils.Md5;

/**
 * Programme Java permettant d'exporter des intervenants (ProviderContacts)
 * d'une base de données Mongo DB locale vers un fichier Excel
 *
 * @author Thierry Baribaud
 * @version 1.03
 */
public class ExpProviderContacts {

    /**
     * mgoDbServerType : prod pour le serveur de production, pre-prod pour le
     * serveur de pré-production. Valeur par défaut : pre-prod.
     */
    private String mgoDbServerType = "pre-prod";

    /**
     * debugMode : fonctionnement du programme en mode debug (true/false).
     * Valeur par défaut : false.
     */
    private static boolean debugMode = false;

    /**
     * testMode : fonctionnement du programme en mode test (true/false). Valeur
     * par défaut : false.
     */
    private static boolean testMode = false;

    /**
     * path : répertoire où sera déposé le fichier des résultats
     */
    private String path = ".";

    /**
     * filename : nom du fichier contenant les résultats
     */
    private String filename = "providerContacts.xlsx";

    /**
     * unum : référence au service d'urgence (identifiant interne)
     */
    private int unum;

    /**
     * clientCompanyUuid : identifiant universel unique du service d'urgence
     */
    private String clientCompanyUuid = null;

    /**
     * Constructeur principal de la classe ExpProviderContacts
     *
     * @param args arguments en ligne de commande
     * @throws GetArgsException en cas d'erreur avec les paramètres en ligne de
     * commande
     * @throws java.io.IOException en cas d'erreur d'entrée/sortie.
     * @throws utils.DBServerException en cas d'erreur avec le serveur de base
     * de données.
     */
    public ExpProviderContacts(String[] args) throws GetArgsException, IOException, DBServerException {
        ApplicationProperties applicationProperties;
        DBServer mgoServer;
        MongoClient mongoClient;
        MongoDatabase mongoDatabase;

        System.out.println("Création d'une instance de ExpProviderContacts ...");

        System.out.println("Analyse des arguments de la ligne de commande ...");
        this.getArgs(args);
        System.out.println("Argument(s) en ligne de commande lus().");

        System.out.println("Lecture des paramètres d'exécution ...");
        applicationProperties = new ApplicationProperties("ExpProviderContacts.prop");
        System.out.println("Paramètres d'exécution lus.");

        System.out.println("Lecture des paramètres du serveur Mongo ...");
        mgoServer = new DBServer(mgoDbServerType, "mgoserver", applicationProperties);
        System.out.println("Paramètres du serveur Mongo lus.");
        if (debugMode) {
            System.out.println(mgoServer);
        }

        if (debugMode) {
            System.out.println(this.toString());
        }

        System.out.println("Ouverture de la connexion au serveur MongoDb : " + mgoServer.getName());
        mongoClient = new MongoClient(mgoServer.getIpAddress(), (int) mgoServer.getPortNumber());

        System.out.println("Connexion à la base de données : " + mgoServer.getDbName());
        mongoDatabase = mongoClient.getDatabase(mgoServer.getDbName());

        System.out.println("Export des données ...");
        exportProviderContactsToExcel(mongoDatabase);

    }

    /**
     * @param args arguments en ligne de commande
     */
    public static void main(String[] args) {
        ExpProviderContacts expProviderContacts;

        System.out.println("Lancement de ExpProviderContacts ...");
        try {
            expProviderContacts = new ExpProviderContacts(args);
        } catch (GetArgsException | IOException | DBServerException exception) {
            Logger.getLogger(ExpProviderContacts.class.getName()).log(Level.SEVERE, null, exception);
//            Logger.getLogger(ExpProviderContacts.class.getName()).log(Level.INFO, null, exception);
        }

        System.out.println("Fin de ExpProviderContacts.");

    }

    /**
     * Récupère les paramètres en ligne de commande
     *
     * @param args arguments en ligne de commande
     */
    private void getArgs(String[] args) throws GetArgsException {
        int i;
        int n;
        int ip1;
        String currentParam;
        String nextParam;

        n = args.length;
        System.out.println("nargs=" + n);
        for (i = 0; i < n; i++) {
            System.out.println("args[" + i + "]=" + args[i]);
        }
        i = 0;
        while (i < n) {
//            System.out.println("args[" + i + "]=" + Args[i]);
            currentParam = args[i];
            ip1 = i + 1;
            nextParam = (ip1 < n) ? args[ip1] : null;
            switch (currentParam) {
                case "-mgodb":
                    if (nextParam != null) {
                        if (nextParam.equals("pre-prod") || nextParam.equals("prod")) {
                            this.mgoDbServerType = nextParam;
                        } else {
                            throw new GetArgsException("ERREUR : Mauvais serveur Mongo : " + nextParam);
                        }
                        i = ip1;
                    } else {
                        throw new GetArgsException("ERREUR : Serveur Mongo non définie");
                    }
                    break;
                case "-path":
                    if (nextParam != null) {
                        this.path = nextParam;
                        i = ip1;
                    } else {
                        throw new GetArgsException("ERREUR : Répertoire non défini");
                    }
                    break;
                case "-o":
                    if (nextParam != null) {
                        this.filename = nextParam;
                        i = ip1;
                    } else {
                        throw new GetArgsException("ERREUR : Fichier non défini");
                    }
                    break;
                case "-u":
                    if (nextParam != null) {
                        try {
                            this.unum = Integer.parseInt(nextParam);
                            i = ip1;
                        } catch (Exception exception) {
                            throw new GetArgsException("L'identifiant du service d'urgence doit être numérique : " + nextParam);
                        }

                    } else {
                        throw new GetArgsException("ERREUR : Identifiant du service d'urgence non défini");
                    }
                    break;
                case "-clientCompany":
                    if (nextParam != null) {
                        this.clientCompanyUuid = nextParam;
                        i = ip1;
                    } else {
                        throw new GetArgsException("ERREUR : Identifiant UUID du service d'urgence non défini");
                    }
                    break;
                case "-d":
                    setDebugMode(true);
                    break;
                case "-t":
                    setTestMode(true);
                    break;
                default:
                    usage();
                    throw new GetArgsException("ERREUR : Mauvais paramètre : " + currentParam);
            }
            i++;
        }

        if (unum > 0) {
            if (clientCompanyUuid != null) {
                System.out.println("unum:" + unum + ", clientCompanyUuid:" + clientCompanyUuid);
                throw new GetArgsException("ERREUR : Veuillez choisir unum ou uuid");
            } else {
                clientCompanyUuid = Md5.encode("u:" + unum);
            }
        }
    }

    /**
     * Affiche le mode d'utilisation du programme.
     */
    public static void usage() {
        System.out.println("Usage : java ExpProviderContacts"
                + " [-mgodb prod|pre-prod]"
                + " [-p path]"
                + " [-o file]"
                + " [-u unum|-clientCompany uuid]"
                + " [-d] [-t]");
    }

    /**
     * @return mgoDbServerType retourne le type de serveur MongoDb
     */
    private String getMgoDbServerType() {
        return (mgoDbServerType);
    }

    /**
     * @param mgoDbServerType définit le type de serveur MongoDb
     */
    private void setMgoDbServerType(String mgoDbServerType) {
        this.mgoDbServerType = mgoDbServerType;
    }

    /**
     * @return debugMode : retourne le mode de fonctionnement debug.
     */
    public boolean getDebugMode() {
        return (debugMode);
    }

    /**
     * @param debugMode : fonctionnement du programme en mode debug
     * (true/false).
     */
    public void setDebugMode(boolean debugMode) {
        ExpProviderContacts.debugMode = debugMode;
    }

    /**
     * @return testMode : retourne le mode de fonctionnement test.
     */
    public boolean getTestMode() {
        return (testMode);
    }

    /**
     * @param testMode : fonctionnement du programme en mode test (true/false).
     */
    public void setTestMode(boolean testMode) {
        ExpProviderContacts.testMode = testMode;
    }

    /**
     * @return retourne répertoire où sera déposé le fichier des résultats
     */
    public String getPath() {
        return path;
    }

    /**
     * @param path définit répertoire où sera déposé le fichier des résultats
     */
    public void setPath(String path) {
        this.path = path;
    }

    /**
     * @return retourne le nom du fichier contenant les résultats
     */
    public String getFilename() {
        return filename;
    }

    /**
     * @param filename définit le nom du fichier contenant les résultats
     */
    public void setFilename(String filename) {
        this.filename = filename;
    }

    /**
     * @return retourne la référence au service d'urgence (identifiant interne)
     */
    public int getUnum() {
        return unum;
    }

    /**
     * @param unum définit la référence au service d'urgence (identifiant
     * interne)
     */
    public void setUnum(int unum) {
        this.unum = unum;
    }

    /**
     * @return retourne l'identifiant universel unique du service d'urgence
     */
    public String getClientCompanyUuid() {
        return clientCompanyUuid;
    }

    /**
     * @param clientCompanyUuid définit l'identifiant universel unique du
     * service d'urgence
     */
    public void setClientCompanyUuid(String clientCompanyUuid) {
        this.clientCompanyUuid = clientCompanyUuid;
    }

    /**
     * Exporte les données dans le fichier Excel
     */
    private void exportProviderContactsToExcel(MongoDatabase mongoDatabase) {
        FileOutputStream out;
        XSSFWorkbook classeur;
        XSSFSheet feuille;
        XSSFRow titre;
        XSSFCell cell;
        ArrayList<XSSFCell> cells;
        XSSFRow ligne;
        XSSFCellStyle cellStyle;
        XSSFCellStyle titleStyle;
        ObjectMapper objectMapper;
        ProviderContact providerContact;
        CreationHelper createHelper;
        XSSFHyperlink link;
        XSSFCellStyle hlinkStyle;
        XSSFFont hlinkFont;
        String lastName;
        String firstName;
        Name name;
        CivilName civilName;
        PoorName poorName;
        BasicDBObject filter;
        ItemAbstract company;
        ItemAbstractWithRefList patrimonies;
        ItemAbstractWithRef patrimony;
        String ref;
        String label;
        int nbPatrimonies;
        int nbRows;
        int i;
        int j;
        short nbColumns;
        short nbColumns2;
        int nbProviderContacts;

        objectMapper = new ObjectMapper();
        filter = new BasicDBObject("company.uid", clientCompanyUuid);
        System.out.println("filter:"+filter);

        MongoCollection<Document> providerContactsCollection = mongoDatabase.getCollection("providerContacts");
        System.out.println(providerContactsCollection.count() + " intervenant(s)");

//      Création d'un classeur Excel
        classeur = new XSSFWorkbook();
        createHelper = classeur.getCreationHelper();
        feuille = classeur.createSheet("Intervenants");

        // Style de cellule avec bordure noire
        cellStyle = classeur.createCellStyle();
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());

        // Style pour le titre
        titleStyle = (XSSFCellStyle) cellStyle.clone();
        titleStyle.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        titleStyle.setFillPattern(FillPatternType.LESS_DOTS);
//        titleStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());

        // Style pour les liens dans les cellules
        hlinkStyle = (XSSFCellStyle) cellStyle.clone();
        hlinkFont = classeur.createFont();
        hlinkFont.setUnderline(XSSFFont.U_SINGLE);
        hlinkFont.setColor(HSSFColor.BLUE.index);
        hlinkStyle.setFont(hlinkFont);

        // Ligne de titre
        nbColumns = 0;
        nbRows = 0;
        titre = feuille.createRow(nbRows++);
        cell = titre.createCell(nbColumns++);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("Nom");

        cell = titre.createCell(nbColumns++);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("Prénom");

        cell = titre.createCell(nbColumns++);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("ID Performance Immo");

        cell = titre.createCell(nbColumns++);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("UID Société");

        cell = titre.createCell(nbColumns++);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("Société");

        nbColumns2 = nbColumns;
        cell = titre.createCell(nbColumns++);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("Référence");

        cell = titre.createCell(nbColumns++);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("Libellé");

        // Lit les intervenants filtrés par société
        MongoCursor<Document> providerContactsCursor
                = providerContactsCollection.find(filter).iterator();
        nbProviderContacts = 1;
        try {
            while (providerContactsCursor.hasNext()) {
                providerContact = objectMapper.readValue(providerContactsCursor.next().toJson(), ProviderContact.class);
                System.out.println(nbProviderContacts++
                        + " name:" + providerContact.getName()
                        + ", label:" + providerContact.getLabel()
                        + ", uid:" + providerContact.getUid()
                        + ", nbRows:" + nbRows);

                ligne = feuille.createRow(nbRows++);

                name = providerContact.getName();
                if (name instanceof CivilName) {
                    civilName = (CivilName) name;
                    lastName = civilName.getLastName();
                    firstName = civilName.getFirstName();
                } else if (name instanceof PoorName) {
                    poorName = (PoorName) name;
                    lastName = poorName.getValue();
                    firstName = " ";
                } else if ((lastName = providerContact.getLabel()) != null) {
                    firstName = " ";
                } else {
                    lastName = name.getClass().getName();
                    firstName = "class";
                }

                cells = new ArrayList<>();
                cell = ligne.createCell(0);
                cell.setCellValue(lastName);
                cell.setCellStyle(cellStyle);
                cells.add(cell);

                cell = ligne.createCell(1);
                cell.setCellValue(firstName);
                cell.setCellStyle(cellStyle);
                cells.add(cell);

                cell = ligne.createCell(2);
                cell.setCellValue(providerContact.getUid());
//                link = (XSSFHyperlink) createHelper.createHyperlink(HyperlinkType.URL);
//                link.setAddress("https://dashboard.performance-immo.com/providerContacts/" + providerContact.getUid());
//                link.setLabel(providerContact.getUid());
//                cell.setHyperlink((XSSFHyperlink) link);
//                cell.setCellStyle(hlinkStyle);
                cell.setCellStyle(cellStyle);
                cells.add(cell);

                company = providerContact.getCompany();
                cell = ligne.createCell(3);
                cell.setCellValue(company.getUid());
                cell.setCellStyle(cellStyle);
                cells.add(cell);

                cell = ligne.createCell(4);
                cell.setCellValue(company.getLabel());
                cell.setCellStyle(cellStyle);
                cells.add(cell);

                patrimonies = providerContact.getPatrimonies();
                if ((nbPatrimonies = patrimonies.size()) > 0) {
                    for (i = 0; i < nbPatrimonies; i++) {
                        if (i > 0) {
                            ligne = feuille.createRow(nbRows++);
                            for (j = 0; j < nbColumns2; j++) {
                                cell = ligne.createCell(j);
                                cell.setCellValue(cells.get(j).getStringCellValue());
                                cell.setCellStyle(cellStyle);
                            }
                        }
                        patrimony = patrimonies.get(i);
//                        ref = patrimony.getRef();
//                        label = patrimony.getLabel();
                        cell = ligne.createCell(5);
                        cell.setCellValue(patrimony.getRef());
                        cell.setCellStyle(cellStyle);

                        cell = ligne.createCell(6);
                        cell.setCellValue(patrimony.getLabel());
                        cell.setCellStyle(cellStyle);
                    }
                } else {
//                    ref = "";
//                    label = "";
                    cell = ligne.createCell(5);
                    cell.setCellValue("");
                    cell.setCellStyle(cellStyle);

                    cell = ligne.createCell(6);
                    cell.setCellValue("");
                    cell.setCellStyle(cellStyle);
                }
            }

            // Ajustement automatique de la largeur des colonnes
            for (int k = 0; k < nbColumns; k++) {
                feuille.autoSizeColumn(k);
            }

            // Format A4 en sortie
            feuille.getPrintSetup().setPaperSize(PaperSize.A4_PAPER);

            // Orientation paysage
            feuille.getPrintSetup().setLandscape(true);

            // Ajustement à une page en largeur
            feuille.setFitToPage(true);
            feuille.getPrintSetup().setFitWidth((short) 1);
            feuille.getPrintSetup().setFitHeight((short) 0);

            // En-tête et pied de page
            Header header = feuille.getHeader();
            header.setLeft("Liste des intervenants Extranet Anstel");
            header.setRight("&F");

            Footer footer = feuille.getFooter();
            footer.setLeft("Documentation confidentielle Anstel");
            footer.setCenter("Page &P / &N");
            footer.setRight("&D");

            // Ligne à répéter en haut de page
            feuille.setRepeatingRows(CellRangeAddress.valueOf("1:1"));

        } catch (IOException ex) {
            Logger.getLogger(ExpProviderContacts.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            providerContactsCursor.close();
        }

        // Enregistrement du classeur dans un fichier
        try {
            out = new FileOutputStream(new File(getPath() + "\\" + getFilename()));
            classeur.write(out);
            out.close();
            System.out.println("Fichier Excel " + filename + " créé dans " + path);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(ExpProviderContacts.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ExpProviderContacts.class.getName()).log(Level.SEVERE, null, ex);
        }

    }

    /**
     * Retourne le contenu de ExpProviderContacts
     *
     * @return retourne le contenu de ExpProviderContacts
     */
    @Override
    public String toString() {
        return "ExpProviderContacts:{"
                + "mgoDbServerType:" + mgoDbServerType
                + ", path:" + path
                + ", file:" + filename
                + ", unum:" + unum
                + ", clientCompanyUuid:" + clientCompanyUuid
                + ", debugMode:" + debugMode
                + ", testMode:" + testMode
                + "}";
    }

}
