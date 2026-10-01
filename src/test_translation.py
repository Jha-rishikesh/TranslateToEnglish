import os
os.environ["ARGOS_CHUNK_TYPE"] = "MINISBD"

import argostranslate.translate

installed_languages = argostranslate.translate.get_installed_languages()

from_lang = next(
    lang for lang in installed_languages if lang.code == "fr"
)

to_lang = next(
    lang for lang in installed_languages if lang.code == "en"
)

translation = from_lang.get_translation(to_lang)

result = translation.translate(
    "Bonjour, je suis très heureux aujourd'hui."
)

print(result)