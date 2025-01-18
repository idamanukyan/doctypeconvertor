package eu.civitalis;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class TemplateGenerator {

    public static void main(String[] args) {

        File inputDir = new File("/Users/civitalis/Documents/ProcessingFolder");

        if (inputDir.exists() && inputDir.isDirectory()) {
            System.out.println("Directory found: " + inputDir.getAbsolutePath());

            File[] docFiles = inputDir.listFiles((dir, name) -> name.toLowerCase().endsWith(".doc"));
            if (docFiles != null && docFiles.length > 0) {
                System.out.println("Found the following .doc files:");
                for (File docFile : docFiles) {
                    System.out.println("  - " + docFile.getName());
                    try {
                        processDocFile(docFile, new File("/Users/civitalis/Documents/ConvertedFiles"));
                    } catch (IOException e) {
                        System.err.println("Error processing file: " + docFile.getName());
                        e.printStackTrace();
                    }
                }
            } else {
                System.out.println("No .doc files found in the directory.");
            }
        } else {
            System.out.println("The specified path is not a valid directory.");
        }
    }

    private static void processDocFile(File docFile, File outputDir) throws IOException {
        if (!outputDir.exists()) {
            if (outputDir.mkdirs()) {
                System.out.println("Output directory created: " + outputDir.getAbsolutePath());
            } else {
                throw new IOException("Failed to create output directory: " + outputDir.getAbsolutePath());
            }
        }

        try (FileInputStream fis = new FileInputStream(docFile);
             XWPFDocument document = new XWPFDocument()) {

            TemplateHelper.addTitle(document);
            TemplateHelper.addTableContent(document);
            TemplateHelper.createSecondTable(document);
            TemplateHelper.createThirdTable(document);
            TemplateHelper.addRestData(document);

            File outputFile = new File(outputDir, docFile.getName().replace(".doc", ".docx"));
            saveDocument(document, outputFile);

            System.out.println("Document converted and saved as: " + outputFile.getAbsolutePath());
        }
    }

    private static void saveDocument(XWPFDocument document, File outputFile) throws IOException {
        try (FileOutputStream out = new FileOutputStream(outputFile)) {
            document.write(out);
        }
    }
}