# conver-text-uppercase
Convert all the text in the file to UPPER CASE.
# PPTX Processing Script

## Overview

This script processes a PowerPoint presentation (`.pptx` file) by reading its contents, converting the text in the slides to uppercase, and saving the modified presentation as a new `.pptx` file. It uses the following libraries: `jszip`, `pptxgenjs`, and `xml2js`.

## Libraries Used

### 1. **jszip**
   - **Description**: A library for creating, reading, and editing `.zip` files in JavaScript. In this script, it's used to read the contents of a `.pptx` file (which is essentially a `.zip` file with XML and other resources).
   - **Installation**: `npm install jszip`

### 2. **pptxgenjs**
   - **Description**: A JavaScript library for creating PowerPoint (.pptx) files, though it is not actively used in the script provided. It may have been included for future extensions or alternative implementations.
   - **Installation**: `npm install pptxgenjs` (not used directly in this code).

### 3. **xml2js**
   - **Description**: A library to parse XML into JavaScript objects and build JavaScript objects back into XML. This is used to parse the slide XML content and modify the text within the slides.
   - **Installation**: `npm install xml2js`

## Technologies

- **Node.js**: The script is written in JavaScript and executed in a Node.js environment.
- **File System (fs)**: Built-in Node.js module to handle file reading and writing.
- **XML**: Used to represent slide data within `.pptx` files.
- **ZIP Archive**: A `.pptx` file is a ZIP archive containing XML files and resources, which is handled by the `jszip` library.

## Code Explanation

### Step-by-step Process:

1. **Reading the PPTX File**:
   - The PowerPoint file is read using `fs.readFileSync(filePath)` to get the binary data.
   - `JSZip.loadAsync(fileData)` loads this binary data into a ZIP archive.

2. **Identifying Slide Files**:
   - The script filters the ZIP file contents to find all files under the path `ppt/slides/slide`, which contain the slide data.

3. **Parsing and Modifying Slide Content**:
   - Each slide's content is parsed using `xml2js.parseStringPromise(content)`, converting the XML content into a JavaScript object.
   - It then searches for shapes with text content (`p:txBody`), iterating over the paragraphs and text elements (`a:r`, `a:t`) and converts the text to uppercase.

4. **Building and Updating the XML**:
   - After modifying the text, `xml2js.Builder()` is used to convert the modified JavaScript object back into an XML string.
   - The ZIP archive is updated with the modified XML content using `zip.file(slideFile, updatedXml)`.

5. **Creating the Updated PPTX File**:
   - The updated ZIP archive is generated using `zip.generateAsync({ type: 'nodebuffer' })`.
   - Finally, the updated content is written to a new `.pptx` file using `fs.writeFileSync(outputFilePath, updatedData)`.

## How to Run the Script

1. **Install Dependencies**:
   Make sure to install the required libraries by running:
   ```bash
   npm install jszip xml2js pptxgenjs
   ```
