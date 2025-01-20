package eu.civitalis;

import org.apache.poi.xwpf.usermodel.*;

import java.util.List;

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
        try {
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            boolean secondPageReached = false;
            for (XWPFParagraph paragraph : paragraphs) {
                String text = paragraph.getText().trim();
                if (text.isEmpty()) {
                    secondPageReached = true;
                    continue;
                }
                if (secondPageReached && text.startsWith("MOD_")) {
                    return text;
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private static String retrieveProzessNameFromDoc(XWPFDocument document){
        try {
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            StringBuilder firstPageText = new StringBuilder();
            for (XWPFParagraph paragraph : paragraphs) {
                String text = paragraph.getText().trim();

                if (!text.isEmpty()) {
                    firstPageText.append(text).append(" ");
                }

                // Stop collecting text if a page break is detected
                if (paragraph.getCTP().isSetPPr() &&
                        paragraph.getCTP().getPPr().isSetPageBreakBefore()) {
                    break;
                }
            }
            String fullText = firstPageText.toString();
            String searchStart = "mapping für";
            String searchEnd = "Stand:";

            int startIndex = fullText.indexOf(searchStart);
            int endIndex = fullText.indexOf(searchEnd);

            if (startIndex != -1 && endIndex != -1 && startIndex < endIndex) {
                return fullText.substring(startIndex + searchStart.length(), endIndex).trim();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        return null;
    }

    private static String retrieveReleasestandFromDoc(XWPFDocument document){
        try {
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            StringBuilder firstPageText = new StringBuilder();

            for (XWPFParagraph paragraph : paragraphs) {
                String text = paragraph.getText().trim();

                if (!text.isEmpty()) {
                    firstPageText.append(text).append(" ");
                }

                if (paragraph.getCTP().isSetPPr() &&
                        paragraph.getCTP().getPPr().isSetPageBreakBefore()) {
                    break;
                }
            }

            String fullText = firstPageText.toString();
            String searchStart = "Stand:";

            int startIndex = fullText.indexOf(searchStart);

            if (startIndex != -1) {
                startIndex += searchStart.length();
                int endIndex = fullText.indexOf(",", startIndex);

                if (endIndex != -1) {
                    return fullText.substring(startIndex, endIndex).trim();
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        return null;

    }

    private static String retrieveAnpassungAnFromDoc(XWPFDocument document){
        try {
            List<XWPFParagraph> paragraphs = document.getParagraphs();

            StringBuilder firstPageText = new StringBuilder();

            for (XWPFParagraph paragraph : paragraphs) {
                String text = paragraph.getText().trim();

                if (!text.isEmpty()) {
                    firstPageText.append(text).append(" ");
                }

                if (paragraph.getCTP().isSetPPr() &&
                        paragraph.getCTP().getPPr().isSetPageBreakBefore()) {
                    break;
                }
            }

            String fullText = firstPageText.toString();
            String searchStart = "Stand:";

            int lineStartIndex = fullText.indexOf(searchStart);

            if (lineStartIndex != -1) {
                String line = fullText.substring(lineStartIndex);

                int commaIndex = line.indexOf(",");
                int letzteIndex = line.indexOf("letzte");

                if (commaIndex != -1 && letzteIndex != -1 && commaIndex < letzteIndex) {
                    return line.substring(commaIndex + 1, letzteIndex).trim();
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        return null;
    }

    private static String retrieveAnpassungWegenFromDoc(XWPFDocument document){
        try {
            List<XWPFParagraph> paragraphs = document.getParagraphs();

            for (XWPFParagraph paragraph : paragraphs) {
                String text = paragraph.getText().trim();

                if (text.contains("Letzte Änderung:")) {
                    int startIndex = text.indexOf("Letzte Änderung:") + "Letzte Änderung:".length();
                    if (startIndex < text.length()) {
                        return text.substring(startIndex).trim();
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        return null;
    }

    private static void retrieveFromDocMainTable(XWPFDocument document){
    }
}

