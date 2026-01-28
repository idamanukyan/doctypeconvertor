package eu.civitalis;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Optional;

/**
 * Helper class for building document templates and extracting data from legacy .doc files.
 */
public class TemplateHelper {

    private static final Logger logger = LoggerFactory.getLogger(TemplateHelper.class);

    // Constants for document markers
    private static final String MARKER_MOD_PREFIX = "MOD_";
    private static final String MARKER_MAPPING_FOR = "mapping für";
    private static final String MARKER_STAND = "Stand:";
    private static final String MARKER_LETZTE = "letzte";
    private static final String MARKER_LETZTE_AENDERUNG = "Letzte Änderung:";

    // Table labels
    private static final String LABEL_PROZESS = "Prozess";
    private static final String LABEL_RELEASESTAND = "Releasestand";
    private static final String LABEL_LETZTE_ANPASSUNG_AN = "Letzte Anpassung an";
    private static final String LABEL_LETZTE_ANPASSUNG_WEGEN = "Letzte Anpassung wegen";
    private static final String LABEL_FILTERBEDINGUNG = "Filterbedingung";
    private static final String LABEL_JOIN_BEDINGUNGEN = "Join-Bedingungen";
    private static final String LABEL_HILFSVARIABLEN = "Hilfsvariablen";
    private static final String LABEL_TABELLENNAME_VIEW = "Tabellenname/View";
    private static final String LABEL_ATTRIBUTNAME = "Attributname";
    private static final String LABEL_LOGIK = "Logik";

    private static final String DOCUMENT_TITLE = "Risikomanagementsystem (RMS) – Code-Dokumentation";

    private final String documentText;
    private final String firstPageText;

    /**
     * Creates a new TemplateHelper by extracting text from a .doc file.
     *
     * @param docFile the legacy .doc file to read
     * @throws IOException if the file cannot be read
     */
    public TemplateHelper(File docFile) throws IOException {
        if (docFile == null) {
            throw new IllegalArgumentException("docFile cannot be null");
        }
        if (!docFile.exists()) {
            throw new IllegalArgumentException("docFile does not exist: " + docFile.getAbsolutePath());
        }
        if (!docFile.getName().toLowerCase().endsWith(".doc")) {
            throw new IllegalArgumentException("File must be a .doc file: " + docFile.getName());
        }

        logger.debug("Reading document: {}", docFile.getName());

        try (FileInputStream fis = new FileInputStream(docFile);
             HWPFDocument hwpfDocument = new HWPFDocument(fis);
             WordExtractor extractor = new WordExtractor(hwpfDocument)) {

            this.documentText = extractor.getText();
            this.firstPageText = extractFirstPageText(extractor);

            logger.debug("Successfully extracted text from document ({} characters)", documentText.length());
        }
    }

    /**
     * Extracts text from the first page of the document.
     */
    private String extractFirstPageText(WordExtractor extractor) {
        String[] paragraphs = extractor.getParagraphText();
        StringBuilder firstPage = new StringBuilder();

        for (String paragraph : paragraphs) {
            String trimmed = paragraph.trim();
            if (!trimmed.isEmpty()) {
                firstPage.append(trimmed).append(" ");
            }
            // Stop at page break indicator (form feed character)
            if (paragraph.contains("\f")) {
                break;
            }
        }

        return firstPage.toString();
    }

    /**
     * Adds the document title to the output document.
     */
    public void addTitle(XWPFDocument document) {
        if (document == null) {
            logger.warn("Cannot add title: document is null");
            return;
        }

        XWPFParagraph titleParagraph = document.createParagraph();
        XWPFRun titleRun = titleParagraph.createRun();
        titleRun.setText(DOCUMENT_TITLE);
        titleRun.setBold(true);
        titleRun.setFontSize(14);

        // Add spacing after title
        document.createParagraph();

        logger.debug("Added document title");
    }

    /**
     * Adds the main metadata table to the output document.
     */
    public void addTableContent(XWPFDocument document) {
        if (document == null) {
            logger.warn("Cannot add table content: document is null");
            return;
        }

        XWPFTable table = document.createTable(5, 2);
        setTableWidth(table);

        fillRow(table.getRow(0), "", retrieveMappingName().orElse(""));
        fillRow(table.getRow(1), LABEL_PROZESS, retrieveProzessName().orElse(""));
        fillRow(table.getRow(2), LABEL_RELEASESTAND, retrieveReleasestand().orElse(""));
        fillRow(table.getRow(3), LABEL_LETZTE_ANPASSUNG_AN, retrieveAnpassungAn().orElse(""));
        fillRow(table.getRow(4), LABEL_LETZTE_ANPASSUNG_WEGEN, retrieveAnpassungWegen().orElse(""));

        // Add spacing after table
        document.createParagraph();

        logger.debug("Added main metadata table");
    }

    /**
     * Creates the second table with filter and join conditions.
     */
    public void createSecondTable(XWPFDocument document) {
        if (document == null) {
            logger.warn("Cannot create second table: document is null");
            return;
        }

        XWPFTable table = document.createTable(4, 2);
        setTableWidth(table);

        fillRow(table.getRow(0), LABEL_FILTERBEDINGUNG, extractFilterCondition().orElse(""));
        fillRow(table.getRow(1), LABEL_JOIN_BEDINGUNGEN, extractJoinConditions().orElse(""));
        fillRow(table.getRow(2), LABEL_HILFSVARIABLEN, extractHelperVariables().orElse(""));
        fillRow(table.getRow(3), LABEL_TABELLENNAME_VIEW, extractTableName().orElse(""));

        // Add spacing after table
        document.createParagraph();

        logger.debug("Added filter/join conditions table");
    }

    /**
     * Creates the third table with attributes and logic.
     */
    public void createThirdTable(XWPFDocument document) {
        if (document == null) {
            logger.warn("Cannot create third table: document is null");
            return;
        }

        XWPFTable table = document.createTable(2, 2);
        setTableWidth(table);

        // Header row
        fillRow(table.getRow(0), LABEL_ATTRIBUTNAME, LABEL_LOGIK);

        // Extract and add attribute data from the main table
        extractMainTableData(table);

        logger.debug("Added attributes table");
    }

    /**
     * Sets a consistent width for tables.
     */
    private void setTableWidth(XWPFTable table) {
        table.setWidth("100%");
    }

    /**
     * Fills a table row with text in both cells.
     */
    private void fillRow(XWPFTableRow row, String cell1Text, String cell2Text) {
        if (row == null) {
            return;
        }

        XWPFTableCell cell1 = row.getCell(0);
        XWPFTableCell cell2 = row.getCell(1);

        if (cell1 != null) {
            cell1.setText(cell1Text != null ? cell1Text : "");
        }
        if (cell2 != null) {
            cell2.setText(cell2Text != null ? cell2Text : "");
        }
    }

    /**
     * Retrieves the mapping name (starts with MOD_) from the document.
     */
    private Optional<String> retrieveMappingName() {
        if (documentText == null || documentText.isEmpty()) {
            return Optional.empty();
        }

        String[] lines = documentText.split("\\r?\\n");
        boolean passedFirstSection = false;

        for (String line : lines) {
            String trimmed = line.trim();
            if (trimmed.isEmpty()) {
                passedFirstSection = true;
                continue;
            }
            if (passedFirstSection && trimmed.startsWith(MARKER_MOD_PREFIX)) {
                logger.debug("Found mapping name: {}", trimmed);
                return Optional.of(trimmed);
            }
        }

        logger.debug("No mapping name found");
        return Optional.empty();
    }

    /**
     * Retrieves the process name from between "mapping für" and "Stand:".
     */
    private Optional<String> retrieveProzessName() {
        if (firstPageText == null || firstPageText.isEmpty()) {
            return Optional.empty();
        }

        int startIndex = firstPageText.indexOf(MARKER_MAPPING_FOR);
        int endIndex = firstPageText.indexOf(MARKER_STAND);

        if (startIndex != -1 && endIndex != -1 && startIndex < endIndex) {
            String prozessName = firstPageText
                    .substring(startIndex + MARKER_MAPPING_FOR.length(), endIndex)
                    .trim();
            logger.debug("Found process name: {}", prozessName);
            return Optional.of(prozessName);
        }

        logger.debug("No process name found");
        return Optional.empty();
    }

    /**
     * Retrieves the release status after "Stand:".
     */
    private Optional<String> retrieveReleasestand() {
        if (firstPageText == null || firstPageText.isEmpty()) {
            return Optional.empty();
        }

        int startIndex = firstPageText.indexOf(MARKER_STAND);
        if (startIndex == -1) {
            return Optional.empty();
        }

        startIndex += MARKER_STAND.length();
        int endIndex = firstPageText.indexOf(",", startIndex);

        if (endIndex != -1) {
            String releasestand = firstPageText.substring(startIndex, endIndex).trim();
            logger.debug("Found release status: {}", releasestand);
            return Optional.of(releasestand);
        }

        logger.debug("No release status found");
        return Optional.empty();
    }

    /**
     * Retrieves the modification date from between comma and "letzte".
     */
    private Optional<String> retrieveAnpassungAn() {
        if (firstPageText == null || firstPageText.isEmpty()) {
            return Optional.empty();
        }

        int lineStartIndex = firstPageText.indexOf(MARKER_STAND);
        if (lineStartIndex == -1) {
            return Optional.empty();
        }

        String line = firstPageText.substring(lineStartIndex);
        int commaIndex = line.indexOf(",");
        int letzteIndex = line.toLowerCase().indexOf(MARKER_LETZTE.toLowerCase());

        if (commaIndex != -1 && letzteIndex != -1 && commaIndex < letzteIndex) {
            String anpassungAn = line.substring(commaIndex + 1, letzteIndex).trim();
            logger.debug("Found modification date: {}", anpassungAn);
            return Optional.of(anpassungAn);
        }

        logger.debug("No modification date found");
        return Optional.empty();
    }

    /**
     * Retrieves the modification reason after "Letzte Änderung:".
     */
    private Optional<String> retrieveAnpassungWegen() {
        if (documentText == null || documentText.isEmpty()) {
            return Optional.empty();
        }

        int startIndex = documentText.indexOf(MARKER_LETZTE_AENDERUNG);
        if (startIndex == -1) {
            return Optional.empty();
        }

        startIndex += MARKER_LETZTE_AENDERUNG.length();

        // Find the end of the line
        int endIndex = documentText.indexOf("\n", startIndex);
        if (endIndex == -1) {
            endIndex = documentText.length();
        }

        String reason = documentText.substring(startIndex, endIndex).trim();
        if (!reason.isEmpty()) {
            logger.debug("Found modification reason: {}", reason);
            return Optional.of(reason);
        }

        logger.debug("No modification reason found");
        return Optional.empty();
    }

    /**
     * Extracts filter condition from document.
     * Placeholder - implement based on actual document structure.
     */
    private Optional<String> extractFilterCondition() {
        // TODO: Implement based on actual document structure
        return extractSectionContent("Filterbedingung");
    }

    /**
     * Extracts join conditions from document.
     * Placeholder - implement based on actual document structure.
     */
    private Optional<String> extractJoinConditions() {
        // TODO: Implement based on actual document structure
        return extractSectionContent("Join-Bedingungen");
    }

    /**
     * Extracts helper variables from document.
     * Placeholder - implement based on actual document structure.
     */
    private Optional<String> extractHelperVariables() {
        // TODO: Implement based on actual document structure
        return extractSectionContent("Hilfsvariablen");
    }

    /**
     * Extracts table name/view from document.
     * Placeholder - implement based on actual document structure.
     */
    private Optional<String> extractTableName() {
        // TODO: Implement based on actual document structure
        return extractSectionContent("Tabellenname");
    }

    /**
     * Generic method to extract content following a section header.
     */
    private Optional<String> extractSectionContent(String sectionName) {
        if (documentText == null || sectionName == null) {
            return Optional.empty();
        }

        int headerIndex = documentText.indexOf(sectionName);
        if (headerIndex == -1) {
            return Optional.empty();
        }

        // Find the content after the header
        int startIndex = headerIndex + sectionName.length();

        // Skip any colon or whitespace
        while (startIndex < documentText.length() &&
                (documentText.charAt(startIndex) == ':' ||
                        Character.isWhitespace(documentText.charAt(startIndex)))) {
            startIndex++;
        }

        // Find end of line or next section
        int endIndex = documentText.indexOf("\n", startIndex);
        if (endIndex == -1) {
            endIndex = documentText.length();
        }

        String content = documentText.substring(startIndex, endIndex).trim();
        return content.isEmpty() ? Optional.empty() : Optional.of(content);
    }

    /**
     * Extracts data from the main table in the document and adds rows to the output table.
     */
    private void extractMainTableData(XWPFTable outputTable) {
        // TODO: Implement table extraction from HWPFDocument
        // This requires parsing the document's table structure
        // For now, we add a placeholder row
        logger.debug("Main table extraction not yet fully implemented");
    }

    /**
     * Gets the full extracted document text.
     */
    public String getDocumentText() {
        return documentText;
    }

    /**
     * Gets the first page text.
     */
    public String getFirstPageText() {
        return firstPageText;
    }
}
