import argostranslate.translate

installed_languages = argostranslate.translate.get_installed_languages()

print("Installed Languages:")

for language in installed_languages:
    print(language.code, "-", language.name)