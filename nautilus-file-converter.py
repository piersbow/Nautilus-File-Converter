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

import gi, subprocess, threading, os

from gi.repository import Nautilus, GObject, Gio, Gtk, GdkPixbuf
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


def change_extension(old_path, new_extension):
    split_ver = old_path.split('.')
    split_ver[-1] = new_extension
    return '.'.join(split_ver)

def find_uri_not_in_use(new_uri, new_uris):
    i = 1
    while os.path.isfile(new_uri) or new_uri in new_uris:
        new_uri = new_uri.split('.')
        new_uri[-2] = new_uri[-2] + '_' + str(i)
        new_uri = '.'.join(new_uri)
        i = i + 1
    return new_uri

def get_archive_handler(archive_uri):
    if is_zipfile(archive_uri):
        return ZipFile
    elif is_rarfile(archive_uri):
        return RarFile
    else:
        raise TypeError


READ_FORMATS = {}
WRITE_FORMATS = {}

if PIL:
    READ_FORMATS['Image'] = ['image/jpeg', 'image/png', 'image/bmp', 'application/postscript', 'image/gif',
                             'image/x-icon', 'image/x-pcx', 'image/x-portable-pixmap', 'image/tiff', 'image/x-xbm',
                             'image/x-xbitmap', 'video/fli', 'image/vnd.fpx', 'image/vnd.net-fpx',
                             'application/octet-stream', 'windows/metafile', 'image/x-xpixmap', 'image/webp']
    WRITE_FORMATS['Image'] = [{'name': 'JPEG', 'mimes': ['image/jpeg']},
                              {'name': 'PNG', 'mimes': ['image/png']},
                              {'name': 'BMP', 'mimes': ['image/bmp']},
                              {'name': 'PDF', 'mimes': ['application/pdf']},
                              {'name': 'GIF', 'mimes': ['image/gif']},
                              {'name': 'ICO', 'mimes': ['image/x-icon']},
                              {'name': 'WebP', 'mimes': ['image/webp']},
                              {'name': 'EPS', 'mimes': ['application/postscript']}]
    if RAR:
        READ_FORMATS['Comic'] = ['application/vnd.comicbook+zip', 'application/vnd.comicbook-rar']
    else:
        READ_FORMATS['Comic'] = ['application/vnd.comicbook+zip']
    WRITE_FORMATS['Comic'] = [{'name': 'PDF', 'mimes': ['application/pdf']}]

if find_executable('ffmpeg'):
    READ_FORMATS['Audio'] = ['audio/mpeg', 'audio/mpeg3', 'video/x-mpeg', 'audio/x-mpeg-3',  # MP3
                             'audio/x-wav', 'audio/wav', 'audio/wave', 'audio/x-pn-wave', 'audio/vnd.wave',  # WAV
                             'audio/x-mpegurl',  # M3U
                             'audio/mp4', 'audio/mp4a-latm', 'audio/mpeg4-generic',  # M4A
                             'audio/x-matroska',  # MKV (audio)
                             'audio/aac', 'audio/aacp', 'audio/3gpp', 'audio/3gpp2',  # ACC
                             'audio/ogg',  # OGG
                             'audio/opus',  # OPUS
                             'audio/flac']  # FLAC
    WRITE_FORMATS['Audio'] = [{'name': 'MP3', 'mimes': ['audio/mpeg', 'audio/mpeg3', 'video/x-mpeg', 'audio/x-mpeg-3']},
                              {'name': 'WAV', 'mimes': ['audio/x-wav', 'audio/wav', 'audio/wave', 'audio/x-pn-wave',
                                                        'audio/vnd.wave']},
                              {'name': 'AAC', 'mimes': ['audio/aac', 'audio/aacp', 'audio/3gpp', 'audio/3gpp2']},
                              {'name': 'FLAC', 'mimes': ['audio/flac']},
                              {'name': 'M4A', 'mimes': ['audio/mp4', 'audio/mp4a-latm', 'audio/mpeg4-generic']},
                              {'name': 'OGG', 'mimes': ['audio/ogg']},
                              {'name': 'OPUS', 'mimes': ['audio/opus']}]

    READ_FORMATS['Video'] = ['video/mp4', 'video/webm', 'video/x-matroska', 'video/avi', 'video/msvideo',
                             'video/x-msvideo']
    WRITE_FORMATS['Video'] = [{'name': 'MP4', 'mimes': ['video/mp4']},
                              {'name': 'WebM', 'mimes': ['video/webm']},
                              {'name': 'MKV', 'mimes': ['video/x-matroska']},
                              {'name': 'AVI', 'mimes': ['video/avi', 'video/msvideo', 'video/x-msvideo']},
                              {'name': 'GIF', 'mimes': ['image/gif']},
                              {'name': 'MP3', 'mimes': ['audio/mpeg3']},
                              {'name': 'WAV', 'mimes': ['audio/x-wav']}]

if find_executable('pandoc'):
    READ_FORMATS['Document'] = ['application/vnd.openxmlformats-officedocument.wordprocessingml.document',  # DOCX
                                'application/vnd.oasis.opendocument.text',  # ODT
                                'application/epub+zip', 'text/plain']
    WRITE_FORMATS['Document'] = [
        {'name': 'Docx', 'mimes': ['application/vnd.openxmlformats-officedocument.wordprocessingml.document']},
        {'name': 'ODT', 'mimes': ['application/vnd.oasis.opendocument.text']},
        {'name': 'PPT', 'mimes': ['application/vnd.ms-powerpoint']},
        {'name': 'EPub', 'mimes': ['application/epub+zip']},
        {'name': 'TXT', 'mimes': ['text/plain']}]
    for converter in DOC_TO_PDF_CONVERTERS:
        if find_executable(converter):
            WRITE_FORMATS['Document'].append({'name': 'PDF', 'mimes': ['application/pdf']})
            break


class ConverterMenu(GObject.GObject, Nautilus.MenuProvider, Nautilus.LocationWidgetProvider):
    def __init__(self):
        self.infobar_hbox = None
        self.infobar = None
        self.processes = []

    def get_file_items(self, window, files):
        all_mimes = []
        for indv_file in files:  # get all mime types in selected files
            all_mimes.append(indv_file.get_mime_type())
        all_mimes = set(all_mimes)  # remove duplicate MIMES

        # find group of read formats which contain all the mime types
        valid_group = None
        for key in READ_FORMATS.keys():
            if all_mimes <= set(READ_FORMATS[key]):
                valid_group = key
                break
        # if no valid groups exit
        if not valid_group:
            return

        # if only one type of file selected, do not convert to same type again
        if len(all_mimes) == 1:
            mime_not_include = list(all_mimes)[0]
        else:
            mime_not_include = None

        # create main menu called 'Convert to..'
        convert_menu = Nautilus.MenuItem(name='ExampleMenuProvider::ConvertTo', label='Convert to...', tip='', icon='')
        submenu1 = Nautilus.Menu()
        convert_menu.set_submenu(submenu1)  # make submenu

        for write_format in WRITE_FORMATS[valid_group]:
            if mime_not_include not in write_format['mimes']:
                sub_menuitem = Nautilus.MenuItem(name='ExampleMenuProvider::' + write_format['name'],
                                                 label=write_format['name'], tip='', icon='')

                sub_menuitem.connect('activate', self.on_click, files, write_format, valid_group)

                submenu1.append_item(sub_menuitem)

        return convert_menu,

    def get_background_items(self, window, file):
        return None

    def get_widget(self, uri, window):
        self.infobar = Gtk.InfoBar()
        self.infobar.set_message_type(Gtk.MessageType.ERROR)

        return self.infobar

    # converting functions

    def on_click(self, menu, files, write_format, valid_group):
        to_process = []
        new_uris = []
        for file in files:  # create list of items to process
            old_uri = file.get_uri()[7:]  # removes the file:// at start of uri
            if file.get_mime_type() not in write_format['mimes']:  # if the current file is not already the right format
                new_uri = change_extension(old_uri, write_format['name'].lower())  # gets ideal new uri
                new_uri = find_uri_not_in_use(new_uri, new_uris)  # makes sure new uri not in use, if so changes it
                new_uris.append(unquote(new_uri))  # adds this to the list of new file names
                to_process.append({'old_uri': unquote(old_uri), 'new_uri': unquote(new_uri)})

        self.__create_loadingbar("Converting, this may take awhile")

        if valid_group == 'Image':
            for item in to_process:
                Image.open(item['old_uri']).convert('RGB').save(item['new_uri'])
        elif valid_group == 'Audio':
            for item in to_process:
                self.processes.append(
                    subprocess.Popen("ffmpeg -i '" + item['old_uri'] + "' '" + item['new_uri'] + "'", shell=True))
        elif valid_group == 'Video':
            for item in to_process:
                self.processes.append(
                    subprocess.Popen("ffmpeg -i '" + item['old_uri'] + "' '" + item['new_uri'] + "'", shell=True))
        elif valid_group == 'Document':
            for item in to_process:
                self.processes.append(
                    subprocess.Popen("pandoc -o '" + item['new_uri'] + "' '" + item['old_uri'] + "'", shell=True))
        elif valid_group == 'Comic':
            for item in to_process:
                with get_archive_handler(item['old_uri'])(item['old_uri'], 'r') as archive:
                    image_list = []
                    for archive_file_name in archive.namelist():
                        archive_file = archive.open(archive_file_name)
                        image_list.append(Image.open(archive_file).convert())
                if image_list:
                    image_list[0].save(item['new_uri'], "PDF", resolution=100.0, save_all=True, append_images=image_list[1:])

        thread = threading.Thread(target=self.remove_bar_when_done)
        thread.start()
        return True

    def remove_bar_when_done(self):
        job_exits = [job.wait() for job in self.processes]
        self.infobar_hbox.destroy()
        self.infobar.hide()

    # progress bar stuff

    def __create_loadingbar(self, message):
        self.infobar.set_show_close_button(False)
        self.infobar_hbox = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL)

        infobar_msg = Gtk.Label(message)
        self.infobar_hbox.pack_start(infobar_msg, False, False, 0)

        self.infobar.get_content_area().pack_start(self.infobar_hbox, True, True, 0)
        self.infobar.show_all()

    def __remove_loadingbar(self):
        self.infobar_hbox.destroy()
        self.infobar.hide()

        return False
