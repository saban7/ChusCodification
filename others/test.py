import time
import requests
import json
from datetime import datetime
import os

# Enable GPU usage for Ollama (if supported)
os.environ["OLLAMA_USE_CUDA"] = "1"

# Constants
MAX_RETRIES = 3
API_URL = "http://localhost:11434/api/generate"

# Messages to send
messages = [
    "Who is the king of Morocco?",
    "What about Spain?",
    "And France?"
]

# Start script
Starting_time = datetime.now()
print(f"\n⏰ Starting time: '{Starting_time}'")

# Loop through messages and send requests to Ollama
for i in range(3):  # Loop from 0 to 2 (three iterations)
    prompt = messages[i]
    
    # Prepare the API request payload
    data_payload = {
        "model": "llama3.3:70b",
        "prompt": prompt,
        "temperature": 0.0,
        "stream": False,  # Ensures response is returned at once
#        "context": ""  # Clears context to prevent retention
    }
    
    for attempt in range(MAX_RETRIES):
        try:
            response = requests.post(API_URL, headers={'Content-Type': 'application/json'}, json=data_payload)
            response.raise_for_status()
            response_json = response.json()
            api_response = response_json.get("response", "").strip()
            
            print(f"\n📝 Message {i+1}: {prompt}")
            print(f"🤖 LLaMA response: {api_response}\n")
            break  # Exit retry loop on success

        except requests.exceptions.RequestException as req_err:
            print(f"❌ API connection error (attempt {attempt+1}): {req_err}")

        except json.JSONDecodeError as json_err:
            print(f"⚠️ JSON decode error (attempt {attempt+1}): {json_err}")

        if attempt < MAX_RETRIES - 1:
            time.sleep(5)  # Wait before retrying
        else:
            print(f"🚨 Max retries reached for message {i+1}. Skipping to next.")

# End script
Ending_time = datetime.now()
print(f"\n✅ Finished at: {Ending_time}")
