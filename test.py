import ollama

response = ollama.chat(
    model="llama3.2",
    messages=[
        {
            "role": "user",
            "content": "In one sentence, explain what an invoice is."
        }
    ]
)

print(response["message"]["content"])