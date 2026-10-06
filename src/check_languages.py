import argostranslate.translate

installed_languages = argostranslate.translate.get_installed_languages()

for lang in installed_languages:
  print(lang.code, "-", lang.name)