package eu.civitalis;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;

/**
 * Main entry point for converting legacy .doc files to .docx format with a template structure.
 */
public class TemplateGenerator {

    private static final Logger logger = LoggerFactory.getLogger(TemplateGenerator.class);

    private static final String DEFAULT_INPUT_DIR = "./input";
    private static final String DEFAULT_OUTPUT_DIR = "./output";

    public static void main(String[] args) {
        String inputPath = DEFAULT_INPUT_DIR;
        String outputPath = DEFAULT_OUTPUT_DIR;

        // Parse command-line arguments
        if (args.length >= 1) {
            if (args[0].equals("--help") || args[0].equals("-h")) {
                printUsage();
                return;
            }
            inputPath = args[0];
        }
        if (args.length >= 2) {
            outputPath = args[1];
        }

        File inputDir = new File(inputPath);
        File outputDir = new File(outputPath);

        logger.info("Document Converter Starting");
        logger.info("Input directory: {}", inputDir.getAbsolutePath());
        logger.info("Output directory: {}", outputDir.getAbsolutePath());

        if (!validateInputDirectory(inputDir)) {
            return;
        }

        if (!ensureOutputDirectory(outputDir)) {
            return;
        }

        processDirectory(inputDir, outputDir);
    }

    private static void printUsage() {
        System.out.println("Document Type Converter - Converts .doc files to .docx with template");
        System.out.println();
        System.out.println("Usage: java -jar docfilered.jar [INPUT_DIR] [OUTPUT_DIR]");
        System.out.println();
        System.out.println("Arguments:");
        System.out.println("  INPUT_DIR   Directory containing .doc files (default: ./input)");
        System.out.println("  OUTPUT_DIR  Directory for converted .docx files (default: ./output)");
        System.out.println();
        System.out.println("Examples:");
        System.out.println("  java -jar docfilered.jar");
        System.out.println("  java -jar docfilered.jar /path/to/docs");
        System.out.println("  java -jar docfilered.jar /path/to/docs /path/to/output");
    }

    private static boolean validateInputDirectory(File inputDir) {
        if (!inputDir.exists()) {
            logger.error("Input directory does not exist: {}", inputDir.getAbsolutePath());
            System.err.println("Error: Input directory does not exist: " + inputDir.getAbsolutePath());
            return false;
        }

        if (!inputDir.isDirectory()) {
            logger.error("Input path is not a directory: {}", inputDir.getAbsolutePath());
            System.err.println("Error: Input path is not a directory: " + inputDir.getAbsolutePath());
            return false;
        }

        if (!inputDir.canRead()) {
            logger.error("Cannot read input directory: {}", inputDir.getAbsolutePath());
            System.err.println("Error: Cannot read input directory: " + inputDir.getAbsolutePath());
            return false;
        }

        return true;
    }

    private static boolean ensureOutputDirectory(File outputDir) {
        if (!outputDir.exists()) {
            logger.info("Creating output directory: {}", outputDir.getAbsolutePath());
            if (!outputDir.mkdirs()) {
                logger.error("Failed to create output directory: {}", outputDir.getAbsolutePath());
                System.err.println("Error: Failed to create output directory: " + outputDir.getAbsolutePath());
                return false;
            }
        }

        if (!outputDir.canWrite()) {
            logger.error("Cannot write to output directory: {}", outputDir.getAbsolutePath());
            System.err.println("Error: Cannot write to output directory: " + outputDir.getAbsolutePath());
            return false;
        }

        return true;
    }

    private static void processDirectory(File inputDir, File outputDir) {
        File[] docFiles = inputDir.listFiles((dir, name) ->
            name.toLowerCase().endsWith(".doc") && !name.startsWith("~$")
        );

        if (docFiles == null || docFiles.length == 0) {
            logger.warn("No .doc files found in directory: {}", inputDir.getAbsolutePath());
            System.out.println("No .doc files found in the input directory.");
            return;
        }

        // Sort files for consistent processing order
        Arrays.sort(docFiles);

        logger.info("Found {} .doc file(s) to process", docFiles.length);
        System.out.println("Found " + docFiles.length + " .doc file(s) to process:");

        int successCount = 0;
        int failCount = 0;

        for (File docFile : docFiles) {
            System.out.println("  Processing: " + docFile.getName());
            try {
                processDocFile(docFile, outputDir);
                successCount++;
                logger.info("Successfully converted: {}", docFile.getName());
            } catch (Exception e) {
                failCount++;
                logger.error("Failed to convert {}: {}", docFile.getName(), e.getMessage(), e);
                System.err.println("    Error: " + e.getMessage());
            }
        }

        // Print summary
        System.out.println();
        System.out.println("Conversion complete:");
        System.out.println("  Successful: " + successCount);
        System.out.println("  Failed: " + failCount);
        logger.info("Conversion complete. Successful: {}, Failed: {}", successCount, failCount);
    }

    private static void processDocFile(File docFile, File outputDir) throws IOException {
        // Create template helper to read the .doc file
        TemplateHelper templateHelper = new TemplateHelper(docFile);

        // Create new .docx document
        try (XWPFDocument document = new XWPFDocument()) {
            // Apply template structure with extracted data
            templateHelper.addTitle(document);
            templateHelper.addTableContent(document);
            templateHelper.createSecondTable(document);
            templateHelper.createThirdTable(document);

            // Generate output filename
            String outputFileName = docFile.getName().replaceFirst("\\.doc$", ".docx");
            File outputFile = new File(outputDir, outputFileName);

            // Save the document
            saveDocument(document, outputFile);

            System.out.println("    Saved: " + outputFile.getName());
        }
    }

    private static void saveDocument(XWPFDocument document, File outputFile) throws IOException {
        try (FileOutputStream out = new FileOutputStream(outputFile)) {
            document.write(out);
        }
        logger.debug("Document saved: {}", outputFile.getAbsolutePath());
    }
}
