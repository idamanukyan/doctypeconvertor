📄 **DoctypeConvertor**

A Java utility for converting legacy .doc (Microsoft Word 97–2003) files into modern .docx format while applying a predefined template structure.

This tool is designed to streamline document migration and enforce consistency across generated Word documents.

🚀 **Features**

🔍 Recursive search for .doc files in a given directory (including subdirectories).

🔄 Automatic conversion of .doc → .docx.

📝 Template injection: adds titles, tables, and placeholders into converted documents.

🧩 Extensible design: includes placeholder methods for extracting and inserting dynamic content from the original .doc files.

**📂 Project Structure**

1. TemplateGenerator

  Handles file operations and conversion logic.

  main(String[] args) – Entry point, scans directories, finds .doc files, and triggers conversion.

  findDocFiles(File dir) – Recursively searches for .doc files.

  convertor(File docFile, File outputDir) – Converts .doc → .docx using TemplateHelper.

  saveDocument(XWPFDocument document, File outputFile) – Saves a converted document.

2. TemplateHelper

  Applies the predefined template structure to generated .docx documents.

  addTitle() – Adds a bold header: “Risikomanagementsystem (RMS) – Code-Dokumentation”.
  
  addTableContent() – Creates the first table with process details.

  createSecondTable() – Creates a second table with filter/join placeholders.

  createThirdTable() – Creates a third table with attributes and logic (future: populate dynamically).

  Placeholder methods (e.g., retrieveMappingNameFromDoc, retrieveFromDocMainTable) for extracting real content from .doc.

**⚙️ Installation & Setup**

# Clone the repository
git clone https://github.com/<your-username>/DoctypeConvertor.git

# Navigate into the project folder
cd DoctypeConvertor

# Build with Maven or Gradle (example with Maven)
mvn clean install

▶️ Usage
# Run the converter (adjust directory path as needed)
java -cp target/DoctypeConvertor-1.0.jar com.example.TemplateGenerator /path/to/docs

The program will:

  Search for .doc files in the provided directory.

  Convert each file into .docx.

  Apply the predefined template (title + tables + placeholders).

  Save results in the output folder.

**Java (Core logic)**

Apache POI (org.apache.poi.xwpf) for Word document manipulation

**🔮 Future Improvements**

  Implement extraction of real data from .doc main tables.

  Populate placeholders with actual content instead of static/dummy values.

  Add configuration for custom templates.

  Enhance error handling and logging.

👩‍💻 Author

Developed by Ida Manukyan
📧 idamyan01@gmail.com
