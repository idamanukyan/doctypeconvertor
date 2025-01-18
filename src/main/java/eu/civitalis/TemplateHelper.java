package eu.civitalis;

import org.apache.poi.xwpf.usermodel.*;

public class TemplateHelper {


    public static void addTitle(XWPFDocument document) {
        XWPFParagraph titleParagraph = document.createParagraph();
        XWPFRun titleRun = titleParagraph.createRun();
        titleRun.setText("Risikomanagementsystem (RMS) – Code-Dokumentation");
        titleRun.setBold(true);
    }


    public static void addTableContent(XWPFDocument document) {
        XWPFTable table = document.createTable(5, 2);

        fillRow(table.getRow(0), "", retrieveMappingNameFromDoc(document));
        fillRow(table.getRow(1), "Prozess", retrieveProzessNameFromDoc(document));
        fillRow(table.getRow(2), "Releasestand", retrieveReleasestandFromDoc(document));
        fillRow(table.getRow(3), "Letzte Anpassung an", retrieveAnpassungAnFromDoc(document));
        fillRow(table.getRow(4), "Letzte Anpassung wegen", retrieveAnpassungWegenFromDoc(document));
    }

    private static void fillRow(XWPFTableRow row, String cell1Text, String cell2Text) {
        row.getCell(0).setText(cell1Text);
        row.getCell(1).setText(cell2Text);
    }

    public static void createSecondTable(XWPFDocument document){

        XWPFTable table = document.createTable(4,2);

        fillRow(table.getRow(0), "Filterbedingung", "");
        fillRow(table.getRow(1), "Join-Bedingungen", "");
        fillRow(table.getRow(2), "Hilfsvariablen", "");
        fillRow(table.getRow(3), "Tabellenname/View", "");
    }

    public static void createThirdTable(XWPFDocument document){

        XWPFTable table = document.createTable(4,2);

        fillRow(table.getRow(0), "Attributname", "Logik");
        retrieveFromDocMainTable(document);
    }

    public static void addRestData(XWPFDocument document){

    }

    private static String retrieveMappingNameFromDoc(XWPFDocument document){
        return "";
    }

    private static String retrieveProzessNameFromDoc(XWPFDocument document){
        return "";
    }

    private static String retrieveReleasestandFromDoc(XWPFDocument document){
        return "";
    }

    private static String retrieveAnpassungAnFromDoc(XWPFDocument document){
        return "";
    }

    private static String retrieveAnpassungWegenFromDoc(XWPFDocument document){
        return "";
    }

    private static void retrieveFromDocMainTable(XWPFDocument document){
    }
}

