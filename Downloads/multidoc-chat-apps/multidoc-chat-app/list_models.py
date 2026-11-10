import os

from dotenv import load_dotenv
from google import genai

load_dotenv()
client = genai.Client(api_key=os.environ["GOOGLE_API_KEY"])
for model in client.models.list():
    if "generateContent" in (model.supported_actions or []):
        print(model.name)