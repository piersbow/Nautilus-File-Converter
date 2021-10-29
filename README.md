# Nautilus File Converter
Nautilus File Converter is an extension for Nautilus that adds file conversion options to a files context menu.

It supports conversion of images, audio, videos and documents.

## Installation:

1. Install python3-nautilus
   - sudo apt install python3-nautilus
1. Download the source file
    - move it to `/usr/share/nautilus-python/extensions/` for a system-wide installation
    - move it to `~/.local/share/nautilus-python/extensions/` for a user install
1. Restart nautilus
    - `nautilus -q`

### Other dependencies

- For video and audio conversion, install ffmpeg
  * `sudo apt install ffmpeg`
- For image and cbz conversion, install pillow
  * `sudo apt install python3-pil`
- For cbr conversion, install rarfile in addition to pillow
  * `sudo apt install python3-rarfile`
- For basic document conversion, install pandoc
  * `sudo apt install pandoc`
- To export as pdf, install texlive
  * `sudo apt install texlive texlive-latex-extra texlive-xetex`
