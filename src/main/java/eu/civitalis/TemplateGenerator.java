package eu.civitalis;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class TemplateGenerator {

    public static void main(String[] args) {
        File inputDir = new File("/Users/civitalis/Downloads/Datenverarbeitung/01_Bemessungsgrundlage/");

        if (inputDir.exists() && inputDir.isDirectory()) {
            System.out.println("Directory found: " + inputDir.getAbsolutePath());

            findDocFiles(inputDir);

            File[] docFiles = inputDir.listFiles((dir, name) -> name.endsWith(".doc"));

            if (docFiles != null && docFiles.length > 0) {
                System.out.println("Found the following .doc files:");
                for (File docFile : docFiles) {
                    System.out.println(docFile.getName());
                }
            } else {
                System.out.println("No .doc files found in the directory.");
            }
        } else {
            System.out.println("The specified path is not a valid directory.");
        }
    }

    private static void findDocFiles(File dir) {
        File[] files = dir.listFiles();

        if (files != null) {
            for (File file : files) {
                if (file.isDirectory()) {
                    findDocFiles(file); // Recursive call
                } else if (file.isFile() && file.getName().endsWith(".doc")) {
                    System.out.println("Found .doc file: " + file.getAbsolutePath());
                }
            }

        }
    }
    private static void convertor(File docFile, File outputDir) throws IOException {

        try (FileInputStream fis = new FileInputStream(docFile)) {
            XWPFDocument document = new XWPFDocument();

            TemplateHelper.addTitle(document);
            TemplateHelper.addTableContent(document);
            TemplateHelper.createSecondTable(document);
            TemplateHelper.createThirdTable(document);

            File outputFile = new File(outputDir, docFile.getName().replace(".doc", ".docx"));

            saveDocument(document, outputFile);

            System.out.println("Document converted and saved as: " + outputFile.getName());
        }
    }

    private static void saveDocument(XWPFDocument document, File outputFile) throws IOException {
        try (FileOutputStream out = new FileOutputStream(outputFile)) {
            document.write(out);
        }
    }
}


