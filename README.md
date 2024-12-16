# **Power Platform Utilities**

## **Overview**

The **Power Platform Utilities** repository provides a collection of custom APIs designed to extend the functionality of Power Platform applications. The first utility added is for **merging PDF files**.

## **Merge PDFs API**

### **Inputs**
- `meaf_pdfs`: An array of **Base64-encoded strings**, where each string represents a PDF document to be merged.

### **Output**
- `meaf_mergedpdf`: A **Base64-encoded string** representing the merged PDF document.

## **Usage Example**

This custom API can be used within Power Automate to merge multiple PDF files. The input is an array of PDF files encoded in Base64, and the output is a single merged PDF, also returned in Base64 format, which can be saved or used in subsequent actions.
