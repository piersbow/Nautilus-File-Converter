# -*- coding: utf-8 -*-
# Nautilus File Converter 1.1.0
# Copyright (C) 2020 Piers Bowater https://bowater.org/projects/nautilus-file-converter
#
# Nautilus File Converter is free software; you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation; either version 3 of the License, or
# (at your option) any later version.
#
# Nautilus File Converter is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with Nautilus File Converter; if not, see http://www.gnu.org/licenses
# for more information.

import os
import signal
import multiprocessing as mp
import gi
import subprocess
from gi.repository import Nautilus, GObject, Gio, Gtk, GLib, GdkPixbuf
from urllib.parse import unquote
from distutils.spawn import find_executable
from zipfile import ZipFile, is_zipfile

try:
    from PIL import Image

    PIL = True
except ImportError:
    PIL = False

try:
    from rarfile import RarFile, is_rarfile

    RAR = True
except ImportError:
    RAR = False

gi.require_version('Nautilus', '3.0')
gi.require_version('Gtk', '3.0')

DOCUMENT_CONVERTER = 'pandoc'
AUDIO_VIDEO_CONVERTER = 'ffmpeg'
DOC_TO_PDF_CONVERTERS = ('pdflatex', 'xelatex', 'lualatex', 'pdfroff')
DOCUMENT_CONVERTER_FLAGS= []

READ_FORMATS = {}
WRITE_FORMATS = {}

if PIL:
    READ_FORMATS['Image'] = {'image/jpeg', 'image/png', 'image/bmp', 'application/postscript', 'image/gif',
                             'image/x-icon', 'image/x-pcx', 'image/x-portable-pixmap', 'image/tiff', 'image/x-xbm',
                             'image/x-xbitmap', 'video/fli', 'image/vnd.fpx', 'image/vnd.net-fpx',
                             'application/octet-stream', 'windows/metafile', 'image/x-xpixmap', 'image/webp'}
    WRITE_FORMATS['Image'] = [{'name': 'JPEG', 'mimes': ['image/jpeg']},
                              {'name': 'PNG', 'mimes': ['image/png']},
                              {'name': 'BMP', 'mimes': ['image/bmp']},
                              {'name': 'PDF', 'mimes': ['application/pdf']},
                              {'name': 'GIF', 'mimes': ['image/gif']},
                              {'name': 'ICO', 'mimes': ['image/x-icon']},
                              {'name': 'WebP', 'mimes': ['image/webp']},
                              {'name': 'EPS', 'mimes': ['application/postscript']}]
    if RAR:
        READ_FORMATS['Comic'] = {'application/vnd.comicbook+zip', 'application/vnd.comicbook-rar'}
    else:
        READ_FORMATS['Comic'] = {'application/vnd.comicbook+zip'}
    WRITE_FORMATS['Comic'] = [{'name': 'PDF', 'mimes': ['application/pdf']}]

if find_executable('ffmpeg'):
    READ_FORMATS['Audio'] = {'audio/mpeg', 'audio/mpeg3', 'video/x-mpeg', 'audio/x-mpeg-3',  # MP3
                             'audio/x-wav', 'audio/wav', 'audio/wave', 'audio/x-pn-wave', 'audio/vnd.wave',  # WAV
                             'audio/x-mpegurl',  # M3U
                             'audio/mp4', 'audio/mp4a-latm', 'audio/mpeg4-generic',  # M4A
                             'audio/x-matroska',  # MKV (audio)
                             'audio/aac', 'audio/aacp', 'audio/3gpp', 'audio/3gpp2',  # ACC
                             'audio/ogg',  # OGG
                             'audio/opus',  # OPUS
                             'audio/flac'}  # FLAC
    WRITE_FORMATS['Audio'] = [{'name': 'MP3', 'mimes': ['audio/mpeg', 'audio/mpeg3', 'video/x-mpeg', 'audio/x-mpeg-3']},
                              {'name': 'WAV', 'mimes': ['audio/x-wav', 'audio/wav', 'audio/wave', 'audio/x-pn-wave',
                                                        'audio/vnd.wave']},
                              {'name': 'AAC', 'mimes': ['audio/aac', 'audio/aacp', 'audio/3gpp', 'audio/3gpp2']},
                              {'name': 'FLAC', 'mimes': ['audio/flac']},
                              {'name': 'M4A', 'mimes': ['audio/mp4', 'audio/mp4a-latm', 'audio/mpeg4-generic']},
                              {'name': 'OGG', 'mimes': ['audio/ogg']},
                              {'name': 'OPUS', 'mimes': ['audio/opus']}]

    READ_FORMATS['Video'] = {'video/mp4', 'video/webm', 'video/x-matroska', 'video/avi', 'video/msvideo',
                             'video/x-msvideo'}
    WRITE_FORMATS['Video'] = [{'name': 'MP4', 'mimes': ['video/mp4']},
                              {'name': 'WebM', 'mimes': ['video/webm']},
                              {'name': 'MKV', 'mimes': ['video/x-matroska']},
                              {'name': 'AVI', 'mimes': ['video/avi', 'video/msvideo', 'video/x-msvideo']},
                              {'name': 'GIF', 'mimes': ['image/gif']},
                              {'name': 'MP3', 'mimes': ['audio/mpeg3']},
                              {'name': 'WAV', 'mimes': ['audio/x-wav']}]

if find_executable('pandoc'):
    READ_FORMATS['Document'] = {'application/vnd.openxmlformats-officedocument.wordprocessingml.document',  # DOCX
                                'application/vnd.oasis.opendocument.text',  # ODT
                                'application/epub+zip', 'text/plain'}
    WRITE_FORMATS['Document'] = [
        {'name': 'DOCX', 'mimes': ['application/vnd.openxmlformats-officedocument.wordprocessingml.document']},
        {'name': 'ODT', 'mimes': ['application/vnd.oasis.opendocument.text']},
        {'name': 'PPT', 'mimes': ['application/vnd.ms-powerpoint']},
        {'name': 'EPub', 'mimes': ['application/epub+zip']},
        {'name': 'TXT', 'mimes': ['text/plain']}]
    for converter in DOC_TO_PDF_CONVERTERS:
        if find_executable(converter):
            WRITE_FORMATS['Document'].append({'name': 'PDF', 'mimes': ['application/pdf']})
            break
    if (find_executable('xelatex')):
        # this flag is needed if the document has fancy characters like emoji
        DOCUMENT_CONVERTER_FLAGS.append('--pdf-engine=xelatex')

def change_extension(old_path, new_extension):
    split_ver = old_path.split('.')
    split_ver[-1] = new_extension
    return '.'.join(split_ver)


def find_uri_not_in_use(new_uri):
    i = 1
    while os.path.isfile(new_uri):
        new_uri = new_uri.split('.')
        new_uri[-2] = new_uri[-2] + '_' + str(i)
        new_uri = '.'.join(new_uri)
        i = i + 1
    return new_uri


def get_archive_handler(archive_uri):
    if is_zipfile(archive_uri):
        return ZipFile
    elif RAR and is_rarfile(archive_uri):
        return RarFile
    else:
        raise TypeError


class ConverterMenu(GObject.GObject, Nautilus.MenuProvider, Nautilus.LocationWidgetProvider):
    def __init__(self):
        super().__init__()
        self.infobar_hbox = None
        self.infobar = None
        self.process = None
        self.other_processes = mp.Queue()

    def get_widget(self, uri, window) -> Gtk.Widget:
        """ This is the method that we have to implement (because we're
        a LocationWidgetProvider) in order to show our infobar.
        """
        self.infobar = Gtk.InfoBar()
        self.infobar.set_message_type(Gtk.MessageType.ERROR)
        self.infobar.connect("response", self.__cb_infobar_response)

        return self.infobar

    def __cb_infobar_response(self, infobar, response):
        """ Callback for the infobar close button.
        """
        if response == Gtk.ResponseType.CLOSE:
            self.infobar_hbox.destroy()
            self.infobar.hide()
            if self.process is not None and self.process.is_alive():
                self.process.terminate()
                self.process.join()
            while not self.other_processes.empty():
                try:
                    os.kill(self.other_processes.get(), signal.SIGTERM)
                except ProcessLookupError:
                    pass

    def get_file_items(self, window, files):
        # create set of each mime type
        all_mimes = set()
        number_mimes = 0
        mime_type = None
        for file in files:
            mime_type = file.get_mime_type()
            if mime_type not in all_mimes:
                number_mimes += 1
                all_mimes.add(mime_type)

        # find group of read formats which contain all the mime types
        valid_group = None
        for key in READ_FORMATS.keys():
            if all_mimes <= READ_FORMATS[key]:
                valid_group = key
                break
        # if no valid groups exit
        if not valid_group:
            return

        # if only one type of file selected, do not convert to same type again
        if number_mimes != 1:
            mime_type = None

        # create main menu called 'Convert to..'
        convert_menu = Nautilus.MenuItem(name='ExampleMenuProvider::ConvertTo', label='Convert to...', tip='', icon='')
        submenu1 = Nautilus.Menu()
        convert_menu.set_submenu(submenu1)  # make submenu

        for write_format in WRITE_FORMATS[valid_group]:
            if mime_type not in write_format['mimes']:
                sub_menuitem = Nautilus.MenuItem(name='ExampleMenuProvider::' + write_format['name'],
                                                 label=write_format['name'], tip='', icon='')

                sub_menuitem.connect('activate', self.on_click, files, write_format, valid_group)

                submenu1.append_item(sub_menuitem)

        return convert_menu,

    def on_click(self, menu, files, write_format, convert_type):
        # create progress bar
        progressbar = self.__create_progressbar()
        progressbar.set_text("Converting")
        progressbar.set_pulse_step = 1.0 / len(files)
        self.infobar.show_all()

        processing_queue = mp.Queue()

        self.process = mp.Process(target=self.__convert_files,
                                  args=(files, convert_type, write_format, processing_queue))
        self.process.daemon = True
        self.process.start()

        GLib.timeout_add(100, self.__update_progressbar, processing_queue, progressbar)

    def __convert_files(self, files, convert_type, write_format, processing_queue):
        for file in files:
            if file.get_mime_type() not in write_format['mimes']:
                old_uri = unquote(file.get_uri()[7:])
                new_uri = find_uri_not_in_use(change_extension(old_uri, write_format['name'].lower()))
                if convert_type == "Video":
                    process = subprocess.Popen(f"exec ffmpeg -i '{old_uri}' '{new_uri}'", shell=True)
                    self.other_processes.put(process.pid)
                    process.wait()
                    self.other_processes.get()

                elif convert_type == "Image":
                    Image.open(old_uri).convert('RGB').save(new_uri)

                elif convert_type == "Audio":
                    process = subprocess.Popen(f"exec ffmpeg -i '{old_uri}' '{new_uri}'", shell=True)
                    self.other_processes.put(process.pid)
                    process.wait()
                    self.other_processes.get()

                elif convert_type == "Document":
                    flags = ''
                    for flag in DOCUMENT_CONVERTER_FLAGS:
                        flags += flag
                    process = subprocess.Popen(f"exec pandoc {flags} '{old_uri}' -o '{new_uri}'", shell=True)

                    self.other_processes.put(process.pid)
                    process.wait()
                    self.other_processes.get()

                elif convert_type == "Comic":
                    with get_archive_handler(old_uri)(old_uri, 'r') as archive:
                        image_list = []
                        for archive_file_name in archive.namelist():
                            if not archive_file_name[-1] == "/":  # not directory
                                archive_file = archive.open(archive_file_name)
                                try:
                                    image_list.append(Image.open(archive_file).convert('RGB'))
                                except IOError:
                                    pass
                    if image_list:
                        image_list[0].save(new_uri, "PDF", resolution=100.0, save_all=True,
                                           append_images=image_list[1:])
                    for image in image_list:
                        image.close()
                processing_queue.put(file.get_name())
        processing_queue.put(None)
        return True

    def __create_progressbar(self) -> Gtk.ProgressBar:
        """ Create the progressbar used to notify that files are currently
        being processed.
        """
        self.infobar.set_show_close_button(True)
        self.infobar.set_message_type(Gtk.MessageType.INFO)
        self.infobar_hbox = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL)

        progressbar = Gtk.ProgressBar()
        self.infobar_hbox.pack_start(progressbar, True, True, 0)
        progressbar.set_show_text(True)

        self.infobar.get_content_area().pack_start(self.infobar_hbox, True, True, 0)
        self.infobar.show_all()

        return progressbar

    def __update_progressbar(self, processing_queue, progressbar) -> bool:
        if not processing_queue.empty():
            fname = processing_queue.get(block=False)
        else:
            return True

        if fname is None:
            self.infobar_hbox.destroy()
            self.infobar.hide()
            if not processing_queue.empty():
                print("Something went wrong, the queue isn't empty :/")
            return False

        progressbar.pulse()
        progressbar.set_text(f"Converting {fname}")
        progressbar.show_all()
        self.infobar_hbox.show_all()
        self.infobar.show_all()
        return True

    def get_background_items(self, window, file):
        return None
