from pinecone import Pinecone as PineconeClient
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Retrieve Pinecone credentials from environment variables
pinecone_api_key = os.getenv("PINECONE_API_KEY")
pinecone_index_name = os.getenv("PINECONE_INDEX_NAME")

# Initialize Pinecone client
client = PineconeClient(api_key=pinecone_api_key)

# Check if the index exists before deleting
if pinecone_index_name in [index.name for index in client.list_indexes()]:
    client.delete_index(pinecone_index_name)
    print(f"Index '{pinecone_index_name}' has been deleted successfully.")
else:
    print(f"Index '{pinecone_index_name}' not found.")