# pywin32
# langchain
# unstructured
from langchain.document_loaders import UnstructuredEmailLoader
from langchain.schema.document import Document


loader = UnstructuredEmailLoader(
    'eml/email_162.eml',
    process_attachments=True,)

data = loader.load()
print(data)