from langchain_openai import ChatOpenAI
import os
from dotenv import load_dotenv
load_dotenv()


chat=ChatOpenAI(
    model=os.getenv("CHAT_MODEL"),
    openai_api_key=os.getenv("CHAT_API_KEY"),
    openai_api_base=os.getenv("CHAT_API_BASE"),
)

