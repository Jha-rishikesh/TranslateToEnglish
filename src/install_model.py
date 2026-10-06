import argostranslate.package

models = [
    "models/translate-de_en-1_3.argosmodel",
    "models/translate-es_en-1_9.argosmodel",
    "models/translate-pt_en-1_9.argosmodel",
    "models/translate-ja_en-1_1.argosmodel",
    "models/translate-zh_en-1_9.argosmodel"
]

for model in models:
    print(f"Installing: {model}")
    argostranslate.package.install_from_path(model)

print("All models installed successfully!")