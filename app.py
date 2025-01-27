from flask import Flask, request, Response, stream_with_context, render_template
import os
from dotenv import load_dotenv
from flask_cors import CORS
from langchain.vectorstores import Pinecone
from langchain_openai import ChatOpenAI, OpenAIEmbeddings
from langchain.prompts import ChatPromptTemplate
from pinecone import Pinecone as PineconeClient
import json
from bs4 import BeautifulSoup

# Initialize Flask app
app = Flask(__name__)
CORS(app)

# Load environment variables
load_dotenv()

# Initialize services
pinecone_index_name = os.getenv("PINECONE_INDEX_NAME")
pinecone_api_key = os.getenv("PINECONE_API_KEY")
pinecone_environment = os.getenv("PINECONE_ENVIRONMENT")
openai_api_key = os.getenv("OPENAI_API_KEY")

embeddings = OpenAIEmbeddings(openai_api_key=openai_api_key)
client = PineconeClient(api_key=pinecone_api_key, environment=pinecone_environment)
index = client.Index(pinecone_index_name)

vectorstore = Pinecone.from_existing_index(
    index_name=pinecone_index_name,
    embedding=embeddings
)
retriever = vectorstore.as_retriever(search_kwargs={"k": 30})

template = """
You are an expert assistant for UNICS Bank. Use only the context provided to answer questions accurately and thoroughly.

FORMATTING RULES:
1. Text Formatting:
   - Use proper sentence spacing
   - Capitalize the first letter of each sentence
   - Use proper punctuation with no extra spaces
   - Always write "UNICS" in all caps
   - Separate paragraphs with a single line break

2. Numbers and Special Characters:
   - Write phone numbers as: (+237) XXXXXXXXX
   - Use hyphens without spaces for ranges
   - Use proper spacing around slashes

3. Lists and Services:
   - Start each new service on a new line
   - Use a hyphen (-) for bullet points
   - Add a space after each bullet point
   - Group similar services under headings

4. Contact Information:
   - Put each piece of contact info on a new line
   - Format addresses consistently
   - Add proper spacing between address components

Context:
{context}

Question: {question}

Answer:
"""

prompt = ChatPromptTemplate.from_template(template)

# Initialize the model globally
model = ChatOpenAI(
    openai_api_key=openai_api_key,
    model="gpt-4o-mini",
    streaming=True,
    temperature=0.7
)

def clean_text(content):
    # Only remove HTML tags while preserving the original text
    soup = BeautifulSoup(content, 'html.parser')
    return soup.get_text()

@app.route('/')
def index():
    return render_template('unics.html')

@app.route('/api/chat', methods=['POST'])
def chat():
    try:
        data = request.get_json()
        user_input = data.get('message', '')

        if not user_input:
            return Response(json.dumps({'error': 'No input provided'}), status=400)

        # Retrieve relevant documents
        docs = retriever.get_relevant_documents(user_input)
        context = "\n".join(doc.page_content[:10000] for doc in docs)

        messages = prompt.format_messages(
            context=context,
            question=user_input
        )

        def generate():
            try:
                response_text = ''
                for chunk in model.stream(messages):
                    if chunk.content:
                        # Accumulate content
                        response_text += chunk.content
                        
                        # Only yield when we have a complete sentence or substantial chunk
                        if ('.' in response_text or '?' in response_text or '!' in response_text or 
                            '\n' in response_text or len(response_text) > 100):
                            yield f"data: {json.dumps({'content': response_text})}\n\n"
                            response_text = ''  # Reset the buffer

                # Send any remaining content in the buffer
                if response_text:
                    yield f"data: {json.dumps({'content': response_text, 'done': True})}\n\n"

            except Exception as e:
                yield f"data: {json.dumps({'error': str(e)})}\n\n"

        return Response(
            stream_with_context(generate()),
            mimetype='text/event-stream'
        )

    except Exception as e:
        return Response(
            json.dumps({'error': str(e)}),
            status=500,
            mimetype='application/json'
        )

if __name__ == '__main__':
    app.run(port=8000, debug=False)