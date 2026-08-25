# readDocx.py

**Code Analysis: Extracting Text from Microsoft Word Documents**

The provided Python code utilizes the `docx` library to interact with Microsoft Word documents (.docx) and extract their text content.

### Functionality

The code defines a single function, `get_text(filename)`, which takes the file path of a Word document as input and returns its extracted text as a string.

Here's a step-by-step breakdown of the function:

1. **Loading the Word Document**: The function uses the `docx.Document()` method to load the Word document from the specified file path.
2. **Iterating over Paragraphs**: The function iterates over each paragraph in the document using a `for` loop.
3. **Extracting Text**: For each paragraph, the function extracts its text content using the `para.text` attribute and appends it to a list called `full_text`.
4. **Joining Text**: After iterating over all paragraphs, the function joins the extracted text using double newline