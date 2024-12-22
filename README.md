# **Power Platform Utilities**

## **Overview**

The **Power Platform Utilities** repository provides a collection of custom APIs designed to extend the functionality of Power Platform applications. The utilities added include:
- **Merging PDF files**
- **Merging Word documents**
- **Exporting data to Excel**

## **Merge PDFs API**

### **Inputs**
- `meaf_pdfs`: An array of **Base64-encoded strings**, where each string represents a PDF document to be merged.

### **Output**
- `meaf_mergedpdf`: A **Base64-encoded string** representing the merged PDF document.

### **Usage Example**
This custom API can be used within Power Automate to merge multiple PDF files. The input is an array of PDF files encoded in Base64, and the output is a single merged PDF, also returned in Base64 format, which can be saved or used in subsequent actions.

## **Merge Word Docs API**

### **Inputs**
- `meaf_docs`: An array of **Base64-encoded strings**, where each string represents a Word document to be merged.

### **Output**
- `meaf_MergedDoc`: A **Base64-encoded string** representing the merged Word document.

### **Usage Example**
This custom API can be used within Power Automate to merge multiple Word documents. The input is an array of Word documents encoded in Base64, and the output is a single merged Word document, also returned in Base64 format, which can be saved or used in subsequent actions.

## **Export to Excel API**

### **Inputs**
- `meaf_DataJson`: A **JSON array** containing the data to be exported to Excel.
- `meaf_MappingJson`: A **JSON object** that defines the mapping between input fields and Excel column headers.

### **Output**
- `meaf_GeneratedExcelBase64`: A **Base64-encoded string** representing the exported Excel file.

### **Usage Example**
This custom API can be used within Power Automate to export data to an Excel file. The input is a JSON array containing the data and a mapping JSON to specify the Excel column headers. The output is a Base64-encoded Excel file, which can be saved or used in subsequent actions.

