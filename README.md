# **Power Platform Utilities**

## **Overview**

The **Power Platform Utilities** repository provides a collection of custom APIs designed to extend the functionality of Power Platform applications. The utilities added include:

- **Merging PDF files**
- **Merging Word documents**
- **Exporting data to Excel**
- **Fuzzy Search in Dataverse** (New!)

---

## **Merge PDFs API**

### **Inputs**
- `meaf_pdfs`: An array of **Base64-encoded strings**, where each string represents a PDF document to be merged.

### **Output**
- `meaf_mergedpdf`: A **Base64-encoded string** representing the merged PDF document.

### **Usage Example**
This custom API can be used within Power Automate to merge multiple PDF files. The input is an array of PDF files encoded in Base64, and the output is a single merged PDF, also returned in Base64 format, which can be saved or used in subsequent actions.

[Example Usage - Merge PDFs](https://github.com/melamriD365/Power-Platform-Utilities/tree/main/Server%20Extensions/Dataverse%20-%20Custom%20Apis/Merge%20Pdfs)

---

## **Merge Word Docs API**

### **Inputs**
- `meaf_docs`: An array of **Base64-encoded strings**, where each string represents a Word document to be merged.
- `meaf_AddPageBreak`: A **boolean** parameter. If set to `true`, a page break is added between each merged document. Defaults to `true` if not provided.

### **Output**
- `meaf_mergeddoc`: A **Base64-encoded string** representing the merged Word document.

### **Usage Example**
This custom API can be used within Power Automate to merge multiple Word documents. The input is an array of Word documents encoded in Base64, and the output is a single merged Word document, also returned in Base64 format, which can be saved or used in subsequent actions.

[Example Usage - Merge Word Docs](https://github.com/melamriD365/Power-Platform-Utilities/tree/main/Server%20Extensions/Dataverse%20-%20Custom%20Apis/Merge%20WordDocs)

---

## **Export to Excel API**

### **Inputs**
- `meaf_DataJson`: A **JSON array** containing the data to be exported to Excel.
- `meaf_MappingJson`: A **JSON object** that defines the mapping between input fields and Excel column headers.

### **Output**
- `meaf_GeneratedExcelBase64`: A **Base64-encoded string** representing the exported Excel file.

### **Usage Example**
This custom API can be used within Power Automate to export data to an Excel file. The input is a JSON array containing the data and a mapping JSON to specify the Excel column headers. The output is a Base64-encoded Excel file, which can be saved or used in subsequent actions.

[Example Usage - Export to Excel](https://github.com/melamriD365/Power-Platform-Utilities/tree/main/Server%20Extensions/Dataverse%20-%20Custom%20Apis/Export%20To%20Excel)

---

## **Fuzzy Search API (Dataverse)**

**Important:** Ensure that Dataverse is properly configured and the target search tables are indexed to fully utilize the capabilities of the Fuzzy Search API.

### **Inputs**
- `meaf_entity`: A **string** representing the logical name of the Dataverse table (entity) to search (e.g., `"account"`).
- `meaf_searchTerm`: A **string** representing the Lucene-style search term (e.g., "MEA FUZION" to search "MEA FUSION").
- `meaf_selectColumns`: A **string[]** listing which columns to retrieve in the search results (e.g., `["name", "createdon"]`).

### **Outputs**
- `meaf_results`: A **JSON string** representing the complete search results, including `Count`, `Value`, `Score`, `Attributes`, etc.
- `meaf_bestscore`: A **double** representing the highest (best) relevance score among all returned records.
- `meaf_bestrecord`: An **EntityReference** pointing to the top-scored record (if any results are found).

### **Usage Example**
1. **Invoke** the API from Power Automate or other clients with a body like:
   ```json
   {
     "meaf_entity": "account",
     "meaf_searchTerm": "MEA FUZION",
     "meaf_selectColumns": ["name"]
   }
2. ** Process the Response:**
Process the Response:
- `meaf_results` is a JSON string containing the matched records and their details.
- `meaf_bestscore` is a numeric score representing the top match's relevance.
- `meaf_bestrecord` is the EntityReference of the top-scoring record if one exists.

This custom API leverages advanced Dataverse search capabilities—such as fuzzy and proximity matching—to locate, score, and return records in a single request.

[Example usage - Fuzzy Search](https://github.com/melamriD365/Power-Platform-Utilities/tree/main/Server%20Extensions/Dataverse%20-%20Custom%20Apis/Fuzzy%20Search%20API%20(Dataverse))
