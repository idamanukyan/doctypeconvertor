package eu.civitalis;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class TemplateGenerator {

    public static void main(String[] args) {

        File inputDir = new File("path/to/your/doc/files");
        File outputDir = new File("converted");
        if (!outputDir.exists()) {
            outputDir.mkdirs();
        }

        File[] docFiles = inputDir.listFiles((dir, name) -> name.endsWith(".doc"));

        if (docFiles != null) {
            for (File docFile : docFiles) {
                try {
                    System.out.println("Converting: " + docFile.getName());
                    convertor(docFile, outputDir);
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        } else {
            System.out.println("No .doc files found in the specified directory.");

        }
    }

    private static void convertor(File docFile, File outputDir) throws IOException {

        try (FileInputStream fis = new FileInputStream(docFile)) {
            XWPFDocument document = new XWPFDocument();

            TemplateHelper.addTitle(document);
            TemplateHelper.addTableContent(document);

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


