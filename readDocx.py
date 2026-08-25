```python
# Import the required library to interact with Microsoft Word documents (.docx)
import docx

# Function to extract text from a Word document
def get_text(filename):
    # Load the Word document
    doc = docx.Document(filename)
    
    # Initialize an empty list to store the text from all paragraphs
    full_text = []
    
    # Iterate over each paragraph in the document
    for para in doc.paragraphs:
        # Append the text of the current paragraph to the full_text list
        full_text.append(para.text)
    
    # Join the text from all paragraphs with double newline characters for readability
    return '\n\n'.join(full_text)
```