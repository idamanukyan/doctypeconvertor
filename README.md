# DoctypeConvertor

A Java utility for converting legacy .doc (Microsoft Word 97-2003) files into modern .docx format while applying a predefined template structure.

This tool is designed to streamline document migration and enforce consistency across generated Word documents.

## Features

- Batch conversion of .doc files to .docx format
- Automatic text extraction from legacy Word documents using Apache POI HWPF
- Template injection: adds titles, tables, and structured content into converted documents
- Command-line interface with configurable input/output directories
- Comprehensive logging with SLF4J
- Input validation and error handling

## Requirements

- Java 23 or higher
- Maven 3.6+

## Installation

```bash
# Clone the repository
git clone https://github.com/idamanukyan/doctypeconvertor.git

# Navigate into the project folder
cd doctypeconvertor

# Build with Maven
mvn clean package
```

## Usage

### Basic Usage

```bash
# Using default directories (./input and ./output)
java -jar target/docfilered-1.0-SNAPSHOT.jar

# Specify input directory only (output defaults to ./output)
java -jar target/docfilered-1.0-SNAPSHOT.jar /path/to/docs

# Specify both input and output directories
java -jar target/docfilered-1.0-SNAPSHOT.jar /path/to/docs /path/to/output

# Show help
java -jar target/docfilered-1.0-SNAPSHOT.jar --help
```

### Example Output

```
Document Converter Starting
Found 3 .doc file(s) to process:
  Processing: document1.doc
    Saved: document1.docx
  Processing: document2.doc
    Saved: document2.docx
  Processing: document3.doc
    Saved: document3.docx

Conversion complete:
  Successful: 3
  Failed: 0
```

## Project Structure

```
src/main/java/eu/civitalis/
├── TemplateGenerator.java    # Main entry point, CLI handling, batch processing
└── TemplateHelper.java       # Document reading and template generation
```

### TemplateGenerator

Handles command-line parsing, file discovery, and orchestrates the conversion process.

- `main(String[] args)` - Entry point with CLI argument parsing
- `processDirectory()` - Finds and processes all .doc files
- `processDocFile()` - Converts a single .doc file to .docx
- `saveDocument()` - Writes the converted document to disk

### TemplateHelper

Reads legacy .doc files and applies the template structure to new .docx documents.

- Constructor reads .doc file using `HWPFDocument` and extracts text
- `addTitle()` - Adds document header
- `addTableContent()` - Creates metadata table with extracted values
- `createSecondTable()` - Creates filter/join conditions table
- `createThirdTable()` - Creates attributes and logic table
- Various `retrieve*()` methods extract specific data from the source document

## Dependencies

| Dependency | Version | Purpose |
|------------|---------|---------|
| Apache POI | 5.2.5 | Word document manipulation (.doc and .docx) |
| Apache Commons IO | 2.15.1 | File I/O utilities |
| SLF4J | 2.0.9 | Logging framework |

## Template Structure

The converter generates .docx files with the following structure:

1. **Title**: "Risikomanagementsystem (RMS) - Code-Dokumentation"
2. **Metadata Table**: Mapping name, process, release status, modification info
3. **Filter Table**: Filter conditions, join conditions, helper variables, table name
4. **Attributes Table**: Attribute names and logic

## Future Improvements

- [ ] Extract table data from source .doc files
- [ ] Add configuration file support for custom templates
- [ ] Support for recursive directory scanning
- [ ] Add unit tests

## Author

Developed by Ida Manukyan
idamyan01@gmail.com
