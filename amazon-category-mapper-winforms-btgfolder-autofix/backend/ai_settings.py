"""
Opcionalni AI fallback.

Normalno ti ovo NE treba jer program prvo koristi:
1) European Browse Node Mapping za FR/IT/ES
2) lokalni BTG folder i similarity logiku za ostale marketplaceove

Ako jednog dana želiš probati AI za retke gdje BTG similarity ne nađe dobar prijedlog:
- upiši API key u OPENAI_API_KEY
- postavi AI_ENABLED = "true"
- restartaj Python backend

AI rezultat se nikad ne sprema kao potvrđen; izlazi kao prijedlog za ručnu provjeru.
"""

AI_ENABLED = "false"
OPENAI_API_KEY = ""
OPENAI_MODEL = "gpt-4.1-mini"
