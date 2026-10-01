import argostranslate.package

package_path = "models/translate-fr_en-1_9.argosmodel"

argostranslate.package.install_from_path(package_path)

print("French -> English model installed successfully")