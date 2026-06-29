"""
API ključevi za prijevod kategorija.

Program radi i bez ovih ključeva, ali tada za NL/PL/IE/SE koristi samo lokalnu BTG sličnost.
Ako želiš bolji prijevod zadnjeg dijela njemačke kategorije, upiši jedan od API ključeva.

Preporuka:
1) DeepL je najjednostavniji za EU jezike.
2) Azure/Google su također OK.
3) OpenAI može služiti kao fallback za prijevod kratkih kategorijskih pojmova.
"""

# Uključi/isključi vanjski prijevod za BTG similarity.
TRANSLATION_ENABLED = True

# Moguće vrijednosti: "deepl", "google", "azure", "openai", "auto".
# auto bira prvi provider za koji je postavljen API ključ.
TRANSLATION_PROVIDER = "auto"

OPENAI_API_KEY = "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"
OPENAI_TRANSLATION_MODEL = "gpt-4.1-mini"
