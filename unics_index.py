# import os
# from dotenv import load_dotenv
# import openai
# from openai import OpenAI
# from langchain.text_splitter import RecursiveCharacterTextSplitter
# from pinecone import Pinecone as PineconeClient, ServerlessSpec
# import time  
# import docx 
# # Load environment variables
# load_dotenv()

# openai_api_key = os.getenv("OPENAI_API_KEY")
# client = OpenAI(api_key=openai_api_key)

# # Retrieve environment variables
# pinecone_index_name = os.getenv("PINECONE_INDEX_NAME")
# pinecone_api_key = os.getenv("PINECONE_API_KEY")
# pinecone_environment = os.getenv("PINECONE_ENVIRONMENT")

# # Function to extract text from .docx file
# def extract_text_from_docx(docx_file_path):
#     doc = docx.Document(docx_file_path)
#     full_text = []
#     for para in doc.paragraphs:
#         full_text.append(para.text)
#     return '\n'.join(full_text)


# # Read and process the .docx file
# docx_file_path = "Orfa AI Chatbot Doc v0-2.docx"  # Replace with the actual path to your .docx file
# docx_text_content = extract_text_from_docx(docx_file_path)

# # Combine the content of both files
# combined_text_content = docx_text_content  # Combine .txt and .docx content


# # Initialize the RecursiveCharacterTextSplitter
# text_splitter = RecursiveCharacterTextSplitter(chunk_size=400, chunk_overlap=150)

# # Split the text into chunks
# chunks = text_splitter.split_text(combined_text_content)

# # Function to get OpenAI embeddings
# def get_openai_embeddings(text_list):
#     embeddings = []
#     batch_size = 5  # Define the batch size based on your needs
#     for i in range(0, len(text_list), batch_size):
#         batch = text_list[i:i + batch_size]
#         try:
#             # Ensure correct usage of OpenAI client
#             response = client.embeddings.create(
#                 input=batch,
#                 model="text-embedding-ada-002"
#             )

#             batch_embeddings = [item.embedding for item in response.data]  # Use dot notation
#             embeddings.extend(batch_embeddings)
#             print(f"Processed batch {i // batch_size + 1} of {len(text_list) // batch_size + 1}")
#         except openai.OpenAIError as e:
#             print(f"OpenAI API error for batch starting at index {i}: {e}")
#             # Optionally implement retry logic here
#     return embeddings

# # Embed the chunks using OpenAI
# embeddings = get_openai_embeddings(chunks)

# # Initialize Pinecone client
# client = PineconeClient(api_key=pinecone_api_key, environment=pinecone_environment)

# # Check if the index already exists, if not, create it
# if pinecone_index_name not in [index.name for index in client.list_indexes()]:
#     client.create_index(
#         pinecone_index_name,
#         dimension=1536,  # The dimension for 'text-embedding-ada-002' is 1536
#         metric='cosine',
#         spec=ServerlessSpec(
#             cloud='aws',
#             region='us-east-1'
#         )
#     )

# # Connect to the created index
# index = client.Index(pinecone_index_name)

# # Prepare and index the data
# docs = []
# for i, (chunk, embedding) in enumerate(zip(chunks, embeddings)):
#     doc = {
#         'id': str(i),
#         'values': embedding,  # OpenAI embeddings are already lists
#         'metadata': {'text': chunk}
#     }
#     docs.append(doc)

# # Function to split documents into smaller batches
# def batch_documents(documents, batch_size):
#     for i in range(0, len(documents), batch_size):
#         yield documents[i:i + batch_size]

# # Define a batch size
# batch_size = 100  # Adjust based on Pinecone's upsert limits

# # Upsert the documents to Pinecone in batches
# for i, batch in enumerate(batch_documents(docs, batch_size), start=1):
#     try:
#         index.upsert(vectors=batch)
#         print(f"Successfully upserted batch {i} of {len(docs) // batch_size + 1}")
#         # Optional: Add delay to simulate more visible progress (remove in production)
#         time.sleep(0.5)
#     except Exception as e:
#         print(f"Error during upsert operation for batch {i}: {e}")
#         # Optionally implement retry logic here

# print(f"Successfully indexed {len(docs)} chunks into Pinecone")

from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
import os
from dotenv import load_dotenv
import openai
from openai import OpenAI
from langchain.text_splitter import RecursiveCharacterTextSplitter
from pinecone import Pinecone as PineconeClient, ServerlessSpec
import time  
import io
import docx
import PyPDF2
from pptx import Presentation  # Added for PowerPoint files

# Load environment variables
load_dotenv()

openai_api_key = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=openai_api_key)

pinecone_index_name = os.getenv("PINECONE_INDEX_NAME")
pinecone_api_key = os.getenv("PINECONE_API_KEY")
pinecone_environment = os.getenv("PINECONE_ENVIRONMENT")
GOOGLE_DRIVE_FOLDER_ID=os.getenv("GOOGLE_DRIVE_FOLDER_ID")

# Google Drive credentials setup
def authenticate_drive():
    creds = Credentials.from_service_account_file('Google_drive.json')
    return build('drive', 'v3', credentials=creds)

# Function to get all files in the specified folder
def get_files_in_folder(service, folder_id):
    query = f"'{folder_id}' in parents"
    files = []
    page_token = None
    while True:
        response = service.files().list(q=query, fields="files(id, name, mimeType), nextPageToken", pageToken=page_token).execute()
        files.extend(response.get('files', []))
        page_token = response.get('nextPageToken')
        if not page_token:
            break
    if not files:
        raise FileNotFoundError(f"No files found in the folder '{folder_id}'")
    return files

# Function to download the file from Google Drive
def download_file_from_drive(service, file_id, mime_type):
    request = service.files().get_media(fileId=file_id)
    file_stream = io.BytesIO()
    from googleapiclient.http import MediaIoBaseDownload
    downloader = MediaIoBaseDownload(file_stream, request)
    
    done = False
    while not done:
        _, done = downloader.next_chunk()

    file_stream.seek(0)
    
    # Return the file stream
    return file_stream

# Extract text from Google Docs (Google Drive native format)
def extract_text_from_google_doc(service, file_id):
    request = service.files().export_media(fileId=file_id, mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    file_stream = io.BytesIO(request.execute())
    doc = docx.Document(file_stream)
    full_text = [para.text for para in doc.paragraphs]
    return '\n'.join(full_text)

# Extract text from .docx file
def extract_text_from_docx(file_stream):
    doc = docx.Document(file_stream)
    full_text = [para.text for para in doc.paragraphs]
    return '\n'.join(full_text)

# Extract text from PDF file
def extract_text_from_pdf(file_stream):
    reader = PyPDF2.PdfReader(file_stream)
    text = ""
    for page in reader.pages:
        text += page.extract_text()
    return text

# Extract text from PowerPoint (.pptx) file
def extract_text_from_pptx(file_stream):
    prs = Presentation(file_stream)
    full_text = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                full_text.append(shape.text)
    return '\n'.join(full_text)

# Authenticate Google Drive
service = authenticate_drive()

# Folder ID for "OrfaAI" folder
google_drive_folder_id = GOOGLE_DRIVE_FOLDER_ID

# Get all files in the "OrfaAI" folder
all_files = get_files_in_folder(service, google_drive_folder_id)

# Process all files found
all_text_content = ""
for file in all_files:
    try:
        file_name = file['name']
        file_id = file['id']
        mime_type = file['mimeType']
        print(f"Processing file: {file_name} (ID: {file_id}, MIME Type: {mime_type})")

        # Google Doc file
        if mime_type == 'application/vnd.google-apps.document':
            docx_text_content = extract_text_from_google_doc(service, file_id)

        # .docx file
        elif mime_type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
            file_stream = download_file_from_drive(service, file_id, mime_type)
            docx_text_content = extract_text_from_docx(file_stream)

        # PDF file
        elif mime_type == 'application/pdf':
            file_stream = download_file_from_drive(service, file_id, mime_type)
            docx_text_content = extract_text_from_pdf(file_stream)

        # PowerPoint file (.pptx)
        elif mime_type == 'application/vnd.openxmlformats-officedocument.presentationml.presentation':
            file_stream = download_file_from_drive(service, file_id, mime_type)
            docx_text_content = extract_text_from_pptx(file_stream)

        else:
            print(f"Unsupported file type for {file_name}")
            continue

        all_text_content += f"\n\n--- FILE: {file_name} ---\n\n" + docx_text_content

    except Exception as e:
        print(f"Error processing '{file['name']}': {e}")

# Process all extracted text
text_splitter = RecursiveCharacterTextSplitter(chunk_size=400, chunk_overlap=150)
chunks = text_splitter.split_text(all_text_content)

# Function to get OpenAI embeddings
def get_openai_embeddings(text_list):
    embeddings = []
    batch_size = 5  
    for i in range(0, len(text_list), batch_size):
        batch = text_list[i:i + batch_size]
        try:
            response = client.embeddings.create(
                input=batch,
                model="text-embedding-ada-002"
            )
            batch_embeddings = [item.embedding for item in response.data]
            embeddings.extend(batch_embeddings)
            print(f"Processed batch {i // batch_size + 1} of {len(text_list) // batch_size + 1}")
        except openai.OpenAIError as e:
            print(f"OpenAI API error for batch starting at index {i}: {e}")
    return embeddings

embeddings = get_openai_embeddings(chunks)

# Initialize Pinecone client
client = PineconeClient(api_key=pinecone_api_key, environment=pinecone_environment)

# Check if the index exists, if not, create it
if pinecone_index_name not in [index.name for index in client.list_indexes()]:
    client.create_index(
        pinecone_index_name,
        dimension=1536,
        metric='cosine',
        spec=ServerlessSpec(cloud='aws', region='us-east-1')
    )

index = client.Index(pinecone_index_name)

# Prepare and index the data
docs = []
for i, (chunk, embedding) in enumerate(zip(chunks, embeddings)):
    doc = {
        'id': str(i),
        'values': embedding,
        'metadata': {'text': chunk}
    }
    docs.append(doc)

# Function to split documents into batches
def batch_documents(documents, batch_size):
    for i in range(0, len(documents), batch_size):
        yield documents[i:i + batch_size]

batch_size = 100

# Upsert documents to Pinecone in batches
for i, batch in enumerate(batch_documents(docs, batch_size), start=1):
    try:
        index.upsert(vectors=batch)
        print(f"Successfully upserted batch {i} of {len(docs) // batch_size + 1}")
        time.sleep(0.5)
    except Exception as e:
        print(f"Error during upsert operation for batch {i}: {e}")

print(f"Successfully indexed {len(docs)} chunks into Pinecone")