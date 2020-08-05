# Nautilus-File-Converter
Nautilus File Converter is an extension for nautilus that adds file conversion options to a files context menu, it is written in Python

It supports conversion of images, audio, videos and documents.

Installation:

install Python3-Nautilus
  sudo apt install python3-nautilus
ffmpeg for video conversion
  sudo apt install ffmpeg
python-pillow for image and cbz conversion
  sudo apt install python3-pillow
additionally, for cbr conversion, install python-rar
  sudo apt install python3-rarfile
pandoc for document conversion
  sudo apt install pandoc
to allow export as pdf, additionally install texlive
  sudo apt install texlive

Then downlaod the source file and move it to /usr/share/nautilus-python/extensions/nautilus-file-converter.py
Finally restart nautilus by running the command 
  nautilus -q
