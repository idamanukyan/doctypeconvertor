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

        fillRow(table.getRow(0), "", "Retrieved from the mapping for ...");
        fillRow(table.getRow(1), "Prozess", "Text after mapping for ...");
        fillRow(table.getRow(2), "Releasestand", "Stand: ... after");
        fillRow(table.getRow(3), "Letzte Anpassung an", "Stand: xxx, ... after");
        fillRow(table.getRow(4), "Letzte Anpassung wegen", "Letzte Anpassung: A. Petroll");
    }

    private static void fillRow(XWPFTableRow row, String cell1Text, String cell2Text) {
        row.getCell(0).setText(cell1Text);
        row.getCell(1).setText(cell2Text);
    }

    private static void retrieveFromDoc(){

    }
}

