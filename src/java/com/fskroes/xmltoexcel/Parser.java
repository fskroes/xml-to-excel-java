package com.fskroes.xmltoexcel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Fernando Silva Kroes on 01/27/17
 */

public class XmlToExcel {
    private static Workbook workbook;
    private static int rowNum;

    private final static int NS_RITNUMMER_COLUMN = 0;
    private final static int NS_VERTREKTIJD_COLUMN = 1;
    private final static int NS_VERTREKDATUM_COLUMN = 2;
    private final static int NS_VERTREKVERTRAGING_COLUMN = 3;
    private final static int NS_VERTREKVERTRAGINGTEKST_COLUMN = 4;
    private final static int NS_EINDBESTEMMING_COLUMN = 5;
    private final static int NS_TREINSOORT_COLUMN = 6;
    private final static int NS_ROUTETEKST_COLUMN = 7;
    private final static int NS_VERVOERDER_COLUMN = 8;
    private final static int NS_VERTREKSPOOR_COLUMN = 9;
    private final static int NS_VERTREKSPOOR_SUFFIX = 10;

    public static void main(String[] args) throws Exception {
        /* Name of the created excel file */
        File file = new File("parsedxmlfiles.xlsx");
        if(!file.exists())
        {
            try {
                initXls();

                /* Directory of the xml files that needs to parsed */
                File[] arrayOfFiles = getXMLFiles(new File("ENTER DIRECTORY HERE"));

                getAndReadXml(arrayOfFiles);

            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        try {
            initXls();

            /* Directory of the xml files that needs to parsed */
            File[] arrayOfFiles = getXMLFiles(new File("ENTER DIRECTORY HERE"));

            getAndReadXml(arrayOfFiles);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    /* Get thru the directory and list all the xml files in a array of Files */
    public static File[] getXMLFiles(File folder) {
        List<File> aList = new ArrayList<File>();

        File[] files = folder.listFiles();
        for (File pf : files) {

            if (pf.isFile() && getFileExtensionName(pf).indexOf("xml") != -1) {
                aList.add(pf);
            }
        }

        return aList.toArray(new File[aList.size()]);
    }

    /* Check if the file extension is .xml */
    public static String getFileExtensionName(File f) {
        if (f.getName().indexOf(".") == -1) {
            return "";
        } else {
            return f.getName().substring(f.getName().length() - 3, f.getName().length());
        }
    }


    private static void getAndReadXml(File[] array) throws Exception {
        System.out.println("getAndReadXml");

        NodeList nList = null;

        for (int a = 0; a < array.length; a++) {
            File xmlFile = new File(array[a].getAbsolutePath());
            System.out.println(xmlFile.toString());

            System.out.println("downloading file from " + xmlFile + " ...");
            System.out.println("downloading finished, parsing...");


            Sheet sheet = workbook.getSheetAt(0);

            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            Document doc = dBuilder.parse(xmlFile);

            nList = doc.getElementsByTagName("ActueleVertrekTijden");
            for (int i = 0; i < nList.getLength(); i++) {
                System.out.println("Processing element " + (i+1) + "/" + nList.getLength());
                Node node = nList.item(i);
                if (node.getNodeType() == Node.ELEMENT_NODE) {
                    Element element = (Element) node;
                    String vertrekkendeTrein = element.getElementsByTagName("VertrekkendeTrein").item(0).getTextContent();


                    NodeList prods = element.getElementsByTagName("VertrekkendeTrein");
                    for (int j = 0; j < prods.getLength(); j++) {
                        Node prod = prods.item(j);
                        if (prod.getNodeType() == Node.ELEMENT_NODE) {
                            Element product = (Element) prod;

                            String VertrekVertraging = "";
                            String VertrekVertragingTekst = "";
                            String EindBestemming = "";
                            String TreinSoort = "";
                            String RouteText = "";
                            String Vervoerder = "";
                            String VertrekSpoor = "";

                            String RitNummer = product.getElementsByTagName("RitNummer").item(0).getTextContent();
                            String VertrekTijd = product.getElementsByTagName("VertrekTijd").item(0).getTextContent();
                            try {
                                VertrekVertraging = product.getElementsByTagName("VertrekVertraging").item(0).getTextContent();
                            } catch (Exception ex) { System.err.println("VertrekVertraging" + ex); }

                            try {
                                VertrekVertragingTekst = product.getElementsByTagName("VertrekVertragingTekst").item(0).getTextContent();
                            } catch (Exception ex) { System.err.println("VertrekVertragingTekst" + ex); }

                            try {
                                EindBestemming = product.getElementsByTagName("EindBestemming").item(0).getTextContent();
                            } catch (Exception ex) { System.err.println("EindBestemming" + ex); }

                            try {
                                TreinSoort = product.getElementsByTagName("TreinSoort").item(0).getTextContent();
                            } catch (Exception ex) { System.err.println("TreinSoort" + ex); }

                            try {
                                RouteText = product.getElementsByTagName("RouteTekst").item(0).getTextContent();
                            } catch (Exception ex) { System.err.println("RouteTekst" + ex); }

                            try {
                                Vervoerder = product.getElementsByTagName("Vervoerder").item(0).getTextContent();
                            } catch (Exception ex) { System.err.println("Vervoerder" + ex); }

                            try {
                                VertrekSpoor = product.getElementsByTagName("VertrekSpoor").item(0).getTextContent();
                            } catch (Exception ex) { System.err.println("VertrekSpoor" + ex); }


                            Row row = sheet.createRow(rowNum++);
                            Cell cell = row.createCell(NS_RITNUMMER_COLUMN);
                            // TODO check if null
                            cell.setCellValue(RitNummer);

                            if (VertrekTijd.contains("T")) {
                                String[] tmpVertrektijd = VertrekTijd.split("T");

                                cell = row.createCell(NS_VERTREKTIJD_COLUMN);
                                cell.setCellValue(tmpVertrektijd[1]); // time, right-side

                                cell = row.createCell(NS_VERTREKDATUM_COLUMN);
                                cell.setCellValue(tmpVertrektijd[0]); //date, left-side
                            }

                            cell = row.createCell(NS_VERTREKVERTRAGING_COLUMN);
                            cell.setCellValue(VertrekVertraging);

                            cell = row.createCell(NS_VERTREKVERTRAGINGTEKST_COLUMN);
                            cell.setCellValue(VertrekVertragingTekst);

                            cell = row.createCell(NS_EINDBESTEMMING_COLUMN);
                            cell.setCellValue(EindBestemming);

                            cell = row.createCell(NS_TREINSOORT_COLUMN);
                            cell.setCellValue(TreinSoort);

                            cell = row.createCell(NS_ROUTETEKST_COLUMN);
                            cell.setCellValue(RouteText);

                            cell = row.createCell(NS_VERVOERDER_COLUMN);
                            cell.setCellValue(Vervoerder);

                            if (VertrekSpoor.contains("a")) {
                                String tmp = VertrekSpoor.replace("a", " ");
                                cell = row.createCell(NS_VERTREKSPOOR_COLUMN);
                                cell.setCellValue(tmp);

                                cell = row.createCell(NS_VERTREKSPOOR_SUFFIX);
                                cell.setCellValue("a");
                            }
                            if (VertrekSpoor.contains("b")) {
                                String tmp = VertrekSpoor.replace("b", " ");
                                cell = row.createCell(NS_VERTREKSPOOR_COLUMN);
                                cell.setCellValue(tmp);

                                cell = row.createCell(NS_VERTREKSPOOR_SUFFIX);
                                cell.setCellValue("b");
                            }
                        }
                    }
                }
            }


        }

        excelWritingwriting(nList);
    }

    private static void excelWritingwriting(NodeList nodes) throws IOException, InvalidFormatException {

        System.out.println("Node : " + nodes);

        /* Enter the directory where you want the Excel file to be saved */
        FileOutputStream fileOut = new FileOutputStream("ENTER DIRECTORY HERE");
        workbook.write(fileOut);
        workbook.close();
        fileOut.close();

        System.out.println("getAndReadXml finished, processed " + nodes.getLength() + " !");
    }

    /*  Creating the layout of your Excel file, the collumns.
     *  The attribute names in the xml needs to match to the excel columns your are going to make
     *  */
    private static void initXls() {
        workbook = new XSSFWorkbook();

        CellStyle style = workbook.createCellStyle();
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        style.setFont(boldFont);
        style.setAlignment(CellStyle.ALIGN_CENTER);

        Sheet sheet = workbook.createSheet();
        rowNum = 0;
        Row row = sheet.createRow(rowNum++);
        Cell cell = row.createCell(NS_RITNUMMER_COLUMN);
        cell.setCellValue("RitNummer");
        cell.setCellStyle(style);

        cell = row.createCell(NS_VERTREKTIJD_COLUMN);
        cell.setCellValue("VertrekTijd");
        cell.setCellStyle(style);

        cell = row.createCell(NS_VERTREKDATUM_COLUMN);
        cell.setCellValue("VertrekDatum");
        cell.setCellStyle(style);

        cell = row.createCell(NS_VERTREKVERTRAGING_COLUMN);
        cell.setCellValue("VertrekVertraging");
        cell.setCellStyle(style);

        cell = row.createCell(NS_VERTREKVERTRAGINGTEKST_COLUMN);
        cell.setCellValue("VertrekVertragingTekst");
        cell.setCellStyle(style);

        cell = row.createCell(NS_EINDBESTEMMING_COLUMN);
        cell.setCellValue("EindBestemming");
        cell.setCellStyle(style);

        cell = row.createCell(NS_TREINSOORT_COLUMN);
        cell.setCellValue("TreinSoort");
        cell.setCellStyle(style);

        cell = row.createCell(NS_ROUTETEKST_COLUMN);
        cell.setCellValue("RouteTekst");
        cell.setCellStyle(style);

        cell = row.createCell(NS_VERVOERDER_COLUMN);
        cell.setCellValue("Vervoerder");
        cell.setCellStyle(style);

        cell = row.createCell(NS_VERTREKSPOOR_COLUMN);
        cell.setCellValue("VertrekSpoor");
        cell.setCellStyle(style);

        cell = row.createCell(NS_VERTREKSPOOR_SUFFIX);
        cell.setCellValue("VertrekSpoortSuffix");
        cell.setCellStyle(style);
    }
}