#!/usr/bin/env python3
# 
# https://matroska.org/technical/tagging.html for details on the official tags
#
from __future__ import print_function, unicode_literals

import sys, os
import json
import pickle
import re
import datetime
import subprocess
import tempfile
import xmltodict
import lxml.etree as ET
from PyQt5 import QtCore, QtGui, QtWidgets, uic
from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QApplication, QFileDialog, QListWidgetItem, QMessageBox, QDialog
from PyQt5.QtGui import QColor, QStandardItemModel, QStandardItem
from PyQt5.QtCore import Qt, QDate, QUrl, QModelIndex, QItemSelectionModel
from tmdbv3api import *
from tmdbv3api.exceptions import TMDbException

import pprint

qt_creator_file = sys.path[0] + "/ui/main_window.ui"
Ui_MainWindow, QtBaseClass = uic.loadUiType(qt_creator_file)

#
# class to handle the XML properties for tagging
#
class Property():
    def __init__(self, name, value):
        self.name = name
        self.value = value
        self.child = None

    def setChild(self, name, value):
        self.child = Property(name, value)
        return 0

    def getChild(self):
        return self.child

    def __repr__(self):
        return f"Property({self.name, self.value})"

    def __hash__(self):
        return hash(self.name + str(self.value))

    def __eq__(self, other):
        return self.__hash__() == other.__hash__()


#
# The MediaFile class, represents a media file.
#
class MediaFile():
    ## Patherns to parse some info from the filename
    tvshow_regex = '(?P<show>^\w.+)(?P<season>[sS]\d{2,3})(?P<episode>[eE]\d{2,3})(?P<episode_title>.*)$'
    yearRx = '([\(\[ \.\-])([1-2][0-9]{3})([\.\-\)\]_,+])'
    sizeRx = '([\s-]+?)([0-9]{3,4}[i|p])'
    sizes = ['480p', '720p', '1080p', '480i', '720i', '1080i']
    '''
    Media Type  Stik
    Normal (Music)  1
    Audiobook   2
    Music Video 6
    Movie       9
    TV Show     10
    Booklet     11
    Ringtone    14
    '''
    media_types = { 0: 'Unknown', 1: 'Music', 2: 'Audiobook', 6: 'Music Video', 9: 'Movie',
                10: 'TV Show', 11: 'Booklet', 14: 'Ringtone'}

    ## Stuff to handle extracting the tag info out of the XML format mkv uses.
    tag_xpath = '/Tags/Tag/Simple/Name[string() = $name]/..'
    ## Tags we are interested in.
    # Multiple instances of these tags may be present.
    multi_tags = [ 'ACTOR' ]
    crew_tags =  [ 'DIRECTOR', 'ASSISTANT_DIRECTOR', 'DIRECTOR_OF_PHOTOGRAPHY', 'WRITER', 'CASTING', 
                   'EXECUTIVE_PRODUCER', 'SCREENPLAY', 'ORIGINAL_MUSIC_COMPOSER', 'ART_DIRECTION',
                    ]
    # There should only be one of these, we will update the values, which allows us to edit the data.
    unique_tags =  ['TITLE', 'SHOW', 'SUMMARY', 'SEASON', 'EPISODE', 'DATE_RELEASED', 'SUBTITLE', 'MEDIA_TYPE',
                   'DESCRIPTION', 'TMDB', 'GENRE', 'CASTING', 'SUBTITLE' ]
    #
    # The tags we are interested in pulling from the file.
    #
    metadata_tags = multi_tags + crew_tags + unique_tags
    def __init__(self, file):
        self.file = file
        self.metadata = {'file_path': os.path.dirname(file), 'file_name': os.path.basename(file),
                         'tags': {'cast': [], 'genres': [], 'crew': []}}
        self.properties = []
        self.changes = False

        # get some sane defaults for the file.
        self.parse_filename(os.path.basename(file))
        # see if we have metadata in the file
        self.analyze_file()

    def __str__(self):
        return "{}".format(
            {k: val for k, val in self.__dict__.items() if not str(hex(id(val))) in str(val)}
            )
    #
    # Some properties we'll have multiples of but with different hash values (eg. Director, Actor etc) 
    # Others we only want one of tags such as title, description, summary etc.
    # this is what we do with this.
    #
    def uniqueProperty(self, property):
        #
        # If the hash value already present (name + value)
        # don't add another, just return.
        #
        if property in self.properties:
            return
        if property.name in self.unique_tags:
            for prop in self.properties:
                if property.name == prop.name:
                    prop.value = property.value
                    return
            self.properties.append(property)
            return
        # 
        # If we got here it's either a unique property that hasn't been added or a non unique property that needs to be 
        # added.
        #
        self.properties.append(property)

    def lowercase_keys(self, obj):
        if isinstance(obj, dict):
          obj = {key.lower(): value for key, value in obj.items()}
          for key, value in obj.items():         
            if isinstance(value, list):
              for idx, item in enumerate(value):
                value[idx] = self.lowercase_keys(item)
            obj[key] = self.lowercase_keys(value)
        return obj

    def parse_year(self, file_name):
        year = None
        yearMatch = re.search(self.yearRx, file_name)
        if yearMatch:
            yearInt = int(yearMatch.group(2))
            if yearInt > 1900 and yearInt < (datetime.date.today().year + 1):
                year = yearInt
                file_name = file_name.replace(yearMatch.group(1) + yearMatch.group(2) + yearMatch.group(3), '')
        return file_name, year

    def parse_size(self, file_name):
        size = None
        sizeMatch = re.search(self.sizeRx, file_name)
        if sizeMatch:
            size = sizeMatch.group(2)
            file_name = file_name.replace(sizeMatch.group(1) + sizeMatch.group(2), '')
        return file_name, size

    # We parse the filename and try and set the metadata to some sane defaults before we attempt to extract the actual
    # metadata from the file.
    def parse_filename(self, file_name):
        metadata = getattr(self, 'metadata')
        tags = metadata['tags']
        # Strip extension.
        file, ext = os.path.splitext(file_name)
        metadata['container'] = ext
        # pick off the year.
        file, year = self.parse_year(file)
        file, size = self.parse_size(file)
        if year:
            tags['year'] = QDate.fromString(str(year), 'yyyy').toString(Qt.ISODate)
        else:
            tags['year'] = str(QDate.currentDate().toString(Qt.ISODate))
        if size:
            tags['size'] = size
        tvshow = re.search(self.tvshow_regex, file)
        if tvshow and tvshow.group('season') and tvshow.group('episode'):
            show = re.sub(r"[\._-]", " ", tvshow.group('show'))
            episode = re.sub(r"[Ee]", '', tvshow.group('episode'))
            season = re.sub(r"[Ss]", '', tvshow.group('season'))
            tags['show'] = " ".join(show.split())
            tags['title'] = re.sub(r"[\._-]", " ", tvshow.group('episode_title'))
            tags['title'] = " ".join(tags['title'].split())
            tags['episode'] = int(episode)
            tags['season'] =  int(season)
            tags['media_type'] =  str(10)
            self.uniqueProperty(Property('SHOW', " ".join(show.split())))
            self.uniqueProperty(Property('SEASON', season))
            self.uniqueProperty(Property('EPISODE', episode))
            self.uniqueProperty(Property('TITLE', tags['title']))
            self.uniqueProperty(Property('MEDIA_TYPE', 10))
        else:
            # We assume a movie here bit could be any type of media. 
            self.uniqueProperty(Property('TITLE', file))
            self.uniqueProperty(Property('MEDIA_TYPE', 9))
            tags['title'] = file
            tags['media_type'] = str(9)
        self.metadata['tags'].update(tags)

    def analyze_file(self):
        _, temp_file = tempfile.mkstemp(suffix='.xml')
        try:
            output = subprocess.run(['mkvextract', self.file, 'tags', '--global-tags', temp_file],
                                    text = True, check = True, capture_output = True, universal_newlines = True)
        except subprocess.CalledProcessError as Err:
            print ("Error encountered {Err}")
            os.remove(temp_file)
            return
        try:
            root = ET.parse(temp_file).getroot()
        except ET.XMLSyntaxError as Err:
            print (f"No tag data in {self.file}")
        else:
            os.remove(temp_file)
            xml_tags = dict()
            xml_tags['cast'] = []
            xml_tags['crew'] = []
            for tag in self.metadata_tags:
                elem = root.xpath(self.tag_xpath, name = tag)
                if elem:
                    for item in elem:
                        for sub_elem in item.getchildren():
                            # We have an actor/character combo
                            if sub_elem.text == 'ACTOR':
                                parent = item.xpath('Simple/String')
                                value  = item.xpath('String')
                                actor = {str(tag): value[0].text}
                                prop_actor = Property('ACTOR', value[0].text)
                                if parent:
                                    if parent[0] is not None:
                                        actor['CHARACTER'] = parent[0].text
                                        prop_actor.setChild('CHARACTER', parent[0].text)
                                xml_tags['cast'].append(actor)
                                self.uniqueProperty(prop_actor)
                                break
                                
                            if sub_elem.tag == 'String':
                                if str(tag) in self.crew_tags:
                                    crew_tag = {'job': str(tag), 'person': sub_elem.text }
                                    xml_tags['crew'].append(crew_tag)
                                else:
                                    self.uniqueProperty(Property(tag, sub_elem.text))
                                    xml_tags[str(tag)] = sub_elem.text
            if 'GENRE' in xml_tags:
                xml_tags['genres'] = self.media_file_unpack_genres(xml_tags['GENRE'])
            if 'TMDB' in xml_tags:
                prefix, tmdb_id = xml_tags['TMDB'].split('/')
                xml_tags['tmdb_id'] = tmdb_id
            self.metadata['tags'].update(self.lowercase_keys(xml_tags))
 
    def media_file_pack_genres(self, tags):
        if tags:
            genres = list()
            for genre in tags:
                genres.append(genre)
            return '|'.join(genres)
        else:
            return None

    def media_file_unpack_genres(self, tag):
        return tag.split('|')

    def get_media_file_metadata(self):
        return getattr(self, 'metadata')

    def get_media_file(self):
        return (self.file)

    def get_media_file_tags(self):
        metadata = getattr(self, 'metadata')
        return (metadata['tags'])

    def GenerateXML(self):
        root = ET.Element('Tags')
        tag = ET.Element('Tag')
        root.append(tag)
        targets = ET.Element('Targets')
        tag.append(targets)
        targetTypeValue = ET.Element('TargetTypeValue')
        targetTypeValue.text = '50'
        targets.append(targetTypeValue)
        for property in set(self.properties):
            simple = ET.Element('Simple')
            name = ET.Element('Name')
            name.text = property.name
            simple.append(name)
            tag.append(simple)
            string = ET.Element('String')
#            print (f"{property.name} {property.value}")
            string.text = str(property.value)
            simple.append(string)
            child = property.getChild()
            if child != None:
                child_simple = ET.Element('Simple')
                simple.append (child_simple)
                name = ET.Element('Name')
                name.text = child.name
                child_simple.append(name)
                string = ET.Element('String')
                string.text = child.value
                child_simple.append(string)
        # Generate the xml file.
        return ET.tostring(root, xml_declaration=True, pretty_print=True, 
                             doctype='<!DOCTYPE Tags SYSTEM "matroskatags.dtd">')

    @staticmethod
    def get_media_types():
        return MediaFile.media_types

##
## This list represents the list of media files we are working on
## 
class MediaFileModel(QtCore.QAbstractListModel):
    def __init__(self, *args, mediafiles=None, **kwargs):
        super(MediaFileModel, self).__init__(*args, **kwargs)
        self.mediafiles = mediafiles or []

    def data(self, index, role):
        mediafile = self.mediafiles[index.row()]
        metadata = mediafile.get_media_file_metadata()
        if role == Qt.DisplayRole:
            return metadata['file_name']
        if role == Qt.ForegroundRole:
            if mediafile.changes:
                return QColor(Qt.darkRed)
            else:
                return QColor(Qt.darkGreen)

    def rowCount(self, index):
        return len(self.mediafiles)

#
# When we have multiple results from a tmdb search present a dialog to pick one.
#
### Search Results dialog.
class SearchResults(QDialog):
    def __init__(self, results):
        super(SearchResults, self).__init__()
        self.w = loadUi(sys.path[0] + '/ui/tmdb_lookup_results.ui', self)
        self.tmdb = TMDb()
        self.tmdb.language = 'en'
        self.tmdb_config = Configuration().info()
        self.search_results = results
        self.ReleaseDate.clear() 
        for res in results:
            if 'name' in res:
                list_item = QListWidgetItem(res['name'])
            elif 'original_title' in res:
                list_item = QListWidgetItem(res['original_title'])
            list_item.setData(Qt.UserRole, res)
            self.ResultsList.addItem(list_item)
        self.ResultsList.itemClicked.connect(self.ResultsListClicked)
        self.ResultsList.currentRowChanged.connect(self.ResultsListCurrentRowChanged)
        self.ResultsList.itemDoubleClicked.connect(self.ResultListPicked)
        # 9 Times out of ten the first result is the one we want, so lets select that automatically
        self.ResultsList.setCurrentRow(0)
        self.w.show()

    def getSelectedResult(self):
        selected = self.ResultsList.item(self.ResultsList.currentRow())
        return selected

    def fillResult(self, result):
        self.Poster.load(QUrl("about:blank"))
        if 'overview' in result:
            self.Summary.setText(result['overview'])
        if 'release_date' in result:
            date = result['release_date']
        else:
            date = result['first_air_date']

        d = QDate.fromString(date, 'yyyy-MM-dd')
        self.ReleaseDate.setDate(d)

        if 'poster_path' in result:
            if result['poster_path'] is not None:
                poster_url = self.tmdb_config['images']['secure_base_url'] + 'w185' + result['poster_path']
                self.Poster.load(QUrl(poster_url))

    def ResultsListClicked(self, item):
        result = item.data(Qt.UserRole)
        self.fillResult(result)

    def ResultsListCurrentRowChanged(self, currentRow):
        selected_item = self.ResultsList.item(currentRow)
        result = selected_item.data(Qt.UserRole)
        self.fillResult(result)

    def ResultListPicked(self):
        self.buttonBox.accepted.emit()

# 
# Main application
#
class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.current_path = os.getcwd() 
        self.setupUi(self)
        self.model = MediaFileModel()
        self.media_file_view.setModel(self.model)
        self.setup_media_types()
        self.setup_tmdb()
        self.setup_cast_model()
        self.setup_crew_model()

        self.media_file_tvshow_frame.setEnabled(False)
        self.media_file_media_types.setEnabled(False)
        self.media_file_metadata_lookup_btn.setEnabled(False)
        ## Signal Connections
        self.media_file_view.selectionModel().selectionChanged.connect(self.media_file_view_row_changed)
        self.media_file_genre_list.itemClicked.connect(self.media_file_genre_list_item_clicked)
        self.media_file_media_types.activated.connect(self.media_file_media_types_activated)
        self.actionOpenFiles.triggered.connect(self.open_files)
        self.actionSaveFile.triggered.connect(self.save_file)
        self.actionCloseFile.triggered.connect(self.close_file)
        self.actionQuit.triggered.connect(self.close)

        self.media_file_metadata_lookup_btn.clicked.connect(self.media_file_metadata_lookup)

    def setup_cast_model(self):
        self.cast_model = QStandardItemModel()
        self.cast_model.setHorizontalHeaderLabels(['Actor', 'Character'])
        self.media_file_cast_view.setModel(self.cast_model)
        self.media_file_cast_view.horizontalHeader().setSectionResizeMode(1)
        self.media_file_cast_view.horizontalHeader().setDefaultAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)

    def setup_crew_model(self):
        self.crew_model = QStandardItemModel()
        self.crew_model.setHorizontalHeaderLabels(['Person', 'Job'])
        self.media_file_crew_view.setModel(self.crew_model)
        self.media_file_crew_view.horizontalHeader().setSectionResizeMode(1)
        self.media_file_crew_view.horizontalHeader().setDefaultAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)

    def setup_tmdb(self):
        self.tmdb = TMDb()
        self.tmdb.language = "en"
        config = Configuration()
        self.tmdb_config = config.info()
        self.tv_genres = self.get_tv_genres()
        self.movie_genres = self.get_movie_genres()

    def get_tv_genres(self):
        genres = self.load_genres('tv_genres')
        if genres:
            return genres
        else:
            tmdb = Genre()
            genres = tmdb.tv_list()
            self.save_genres('tv_genres', genres)
            return genres

    def get_movie_genres(self):
        genres = self.load_genres('movie_genres')
        if genres:
            return genres
        else:            
            tmdb = Genre()
            genres = tmdb.movie_list()
            self.save_genres('movie_genres', genres)
            return genres 

    def load_genres(self, genre):
        try:
            with open(f"{sys.path[0]}/{genre}.pkl", 'rb') as f:
                return pickle.load(f)
        except:
                return None

    def save_genres(self, genre, genres):
        with open(f"{sys.path[0]}/{genre}.pkl", 'wb') as f:
            pickle.dump(genres, f)

    def setup_media_types(self):
        media_types = MediaFile.get_media_types()
        for key, value in media_types.items():
            self.media_file_media_types.addItem(value, key)

    # Update TV show group
    def update_tvshow(self, metadata):
        tags = metadata['tags']
        self.media_file_tvshow.setText(tags['show'])
        self.media_file_tvshow_season.setValue(int(tags['season']))
        self.media_file_tvshow_episode.setValue(int(tags['episode']))
        if 'summary' in tags:
            self.media_file_tvshow_summary.setPlainText(tags['summary'])
        else:
            self.media_file_tvshow_summary.setPlainText('')
        if 'title' in tags:
            self.media_file_title.setText(tags['title'])
        else:
            self.media_file_title.setText('')
        self.media_file_title_label.setText('Episode Title')

    # Clear tvshow group
    def clear_tvshow(self):
        self.media_file_tvshow.setText("")
        self.media_file_tvshow_season.setValue(0)
        self.media_file_tvshow_episode.setValue(0)
        self.media_file_tvshow_summary.setPlainText('')
        self.media_file_title.setText('')
        self.media_file_title_label.setText("Title")

    # Update all the metadata widgets
    def update_metadata_cast_display(self, cast):
        self.cast_model.removeRows(0, self.cast_model.rowCount())
        for cast_member in cast:
            row = (QStandardItem(cast_member['actor']), QStandardItem(cast_member['character']))
            self.cast_model.appendRow(row)
            self.cast_model.layoutChanged.emit()

    def update_metadata_crew_display(self, crew):
        self.crew_model.removeRows(0, self.crew_model.rowCount())
        for crew_member in crew:
            # Matroska wants tags stored with underscores instead of spaces and in upper case, this will reverse it for 
            # display purposes.
            job = ' '.join(crew_member['job'].split('_')).title()
            row = (QStandardItem(crew_member['person']), QStandardItem(job))
            self.crew_model.appendRow(row)
            self.crew_model.layoutChanged.emit()

    def update_metadata_display(self, metadata):
        tags = metadata['tags']
        self.media_file_media_types.setCurrentIndex(self.media_file_media_types.findData(int(tags['media_type'])))
        self.media_file_file_name.setText(metadata['file_name'])
        if int(tags['media_type']) == 10:
            self.media_file_tvshow_frame.setEnabled(True)
            self.update_tvshow(metadata)
        else:
            self.clear_tvshow()
            self.media_file_tvshow_frame.setEnabled(False)
            self.media_file_title.setText(tags['title'])
        if 'date_released' in tags:
            d = QDate.fromString(tags['date_released'], 'yyyy-MM-dd')
        else:
            d = QDate.fromString(tags['year'], 'yyyy-MM-dd')

        if 'tmdb_id' in tags:
            self.media_file_metadata_id.setText(str(tags['tmdb_id']))
        else:
            self.media_file_metadata_id.setText('')

        self.media_file_release_date.setDate(d)
        if 'description' in tags:
            self.media_file_description.setPlainText(tags['description'])
        else:
            self.media_file_description.setPlainText('')
        self.media_file_update_genres(metadata)
        if 'genre' in tags:
            self.media_file_genre_tag.setText(tags['genre'])
        else:
            self.media_file_genre_tag.setText('')
        if 'title' in tags:
            self.media_file_title.setText(tags['title'])
        self.update_metadata_cast_display(metadata['tags']['cast'])
        self.update_metadata_crew_display(metadata['tags']['crew'])

    def media_file_update_genres(self, metadata):
        self.media_file_genre_list.clear()
        if metadata['tags']['media_type'] == 10:
            genres = self.tv_genres
        else:
            genres = self.movie_genres
        for genre in genres:
            item = QListWidgetItem(genre['name'])
            self.media_file_genre_list.addItem(item)
        for genre in metadata['tags']['genres']:
            matches = self.media_file_genre_list.findItems(genre, Qt.MatchContains)
            for match in matches:
                match.setSelected(True)

    #
    # tag handling
    #
    def media_file_fill_cast_tags(self, file, cast):
        cast_tags = []
        for cast_member in cast:
            actor = Property('ACTOR', cast_member['name'])
            actor.setChild('CHARACTER', cast_member['character'])
            file.uniqueProperty(actor)
            cast_tags.append({'actor': cast_member['name'], 'character': cast_member['character']})
        return cast_tags

    def media_file_fill_crew_tags(self, file, crew):
        crew_tags = []
        for crew_member in crew:
            crew_tags.append({'job': crew_member['job'], 'person': crew_member['name']})         
            tag = '_'.join(crew_member['job'].split(' ')).upper()
            file.uniqueProperty(Property(tag, crew_member['name']))
        return crew_tags

    def media_file_fill_show_tags(self, show):
        index = self.media_file_view.selectionModel().currentIndex()
        file  = self.model.mediafiles[index.row()]
        tags = file.metadata['tags']
        tags['show'] = show['name']
        tags['summary'] = show['overview']
        tags['tmdb_id'] = show['id']
        file.uniqueProperty(Property('SUMMARY', show['overview']))
        file.uniqueProperty(Property('SHOW', show['name']))
        genres = []
        for genre in show['genres']:
            genres.append(genre['name'])
        tags['genres'] = genres 
        tags['genre'] = '|'.join(genres)
        tags['tmdb'] = f"tv/{show['id']}"
        file.uniqueProperty(Property('GENRE', '|'.join(genres)))
        file.uniqueProperty(Property('TMDB', f"tv/{show['id']}"))
        episode = Episode()
        episode_details = episode.details(show['id'], tags['season'], tags['episode'], append_to_response="credits")
        tags['title'] = episode_details['name']
        tags['description'] = episode_details['overview']
        tags['date_released'] = episode_details['air_date']
        tags['cast'] = self.media_file_fill_cast_tags(file, episode_details['credits']['cast'])
        tags['crew'] = self.media_file_fill_crew_tags(file, episode_details['credits']['crew'])
        file.uniqueProperty(Property('TITLE', episode_details['name']))
        file.uniqueProperty(Property('DESCRIPTION', episode_details['overview']))
        file.uniqueProperty(Property('DATE_RELEASED', episode_details['air_date']))
        file.changes = True
        self.update_metadata_display(file.metadata)

    def media_file_fill_movie_tags(self, movie):
        index = self.media_file_view.selectionModel().currentIndex()
        file  = self.model.mediafiles[index.row()]
        tags  = file.metadata['tags']
        tags['tmdb_id'] = movie['id']
        tags['tmdb'] = f"movie/{movie['id']}"
        tags['title'] = movie['title']
        tags['description'] = movie['overview']
        file.uniqueProperty(Property('TMDB', f"movie/{movie['id']}"))
        file.uniqueProperty(Property('TITLE', movie['title']))
        file.uniqueProperty(Property('DESCRIPTION', movie['overview']))
        genres = []
        for genre in movie['genres']:
            genres.append(genre['name'])
        tags['genres'] = genres 
        tags['genre'] = '|'.join(genres)
        tags['date_released'] = movie['release_date']
        file.uniqueProperty(Property('DATE_RELEASED', movie['release_date']))
        file.uniqueProperty(Property('GENRE', '|'.join(genres)))
        tags['cast'] = self.media_file_fill_cast_tags(file, movie['credits']['cast'])
        tags['crew'] = self.media_file_fill_crew_tags(file, movie['credits']['crew'])
        file.changes = True
        pprint.pp(set(file.properties))
        self.update_metadata_display(file.metadata)

    #
    # Metadata searchs/functions
    #
    def media_file_lookup_tvshow(self):
        term = self.media_file_tvshow.text()
        if term:
            tv = TV()
            results = tv.search(term)
            if results['total_results'] == 1:
                tmdb_id = results[0]['id']
                show = tv.details(tmdb_id)
                self.media_file_fill_show_tags(show)
            else:
                self.resultsDialog = SearchResults(results)
                self.resultsDialog.buttonBox.accepted.connect(self.media_file_selected_show)

    def media_file_selected_show(self):
        item = self.resultsDialog.getSelectedResult()
        tmdb_id = item.data(Qt.UserRole)['id']
        tv = TV()
        show = tv.details(tmdb_id)
        self.media_file_fill_show_tags(show)

    def media_file_lookup_movie(self):
        term = self.media_file_title.text()
        if term:
            search = Search()
            try:
                results = search.movies(term, adult = True)
            except TMDbException:
                print ("Movie not found")
                return
            if results['total_results'] == 1:
                movie = Movie()
                details = movie.details(results[0]['id'], append_to_response='credits')
                self.media_file_fill_movie_tags(details)
            else:
                self.resultsDialog = SearchResults(results)
                self.resultsDialog.buttonBox.accepted.connect(self.media_file_selected_movie)

    def media_file_selected_movie(self):
        item = self.resultsDialog.getSelectedResult()
        tmdb_id = item.data(Qt.UserRole)['id']
        movie = Movie()
        details = movie.details(tmdb_id, append_to_response = 'credits')
        self.media_file_fill_movie_tags(details)

#
# Signals
#
    def media_file_view_row_changed(self, selected, deslected):
        indexes = selected.indexes()
        if indexes:
            index = indexes[0]
            file = self.model.mediafiles[index.row()]
            self.media_file_file_path.setText(file.metadata['file_path'])
            self.media_file_media_types.setEnabled(True)
            self.media_file_metadata_lookup_btn.setEnabled(True)        
            self.update_metadata_display(file.get_media_file_metadata())

    def media_file_media_types_activated(self, index):
        media_type = int(self.media_file_media_types.itemData(index))
        index = media_file_view.selectionModel().curentIndex()
        file  = self.model.mediafiles[index.row()]
        if media_type == 10:
            self.media_file_tvshow_frame.setEnabled(True)
        else:
            self.media_file_tvshow_frame.setEnabled(False)
        file.metadata['tags']['media_type'] = media_type
        file.metadata['changed'] = True

    def media_file_genre_list_item_clicked(self, item):
        indexes = self.media_file_view.selectionModel().selectedIndexes()
        if indexes:
            index = indexes[0]
            file  = self.model.mediafiles[index.row()]
            genres = file.metadata['tags']['genres']
            if item.isSelected():
                genres.append(item.text())
            else:
                genres.remove(item.text())
            file.metadata['tags']['genre'] = '|'.join(genres)
            file.metadata['changed'] = True
            self.media_file_genre_tag.setText(file.metadata['tags']['genre'])
        # nothing selected


    def media_file_metadata_lookup (self):
        media_type_name = self.media_file_media_types.currentText()
        media_type = self.media_file_media_types.itemData(self.media_file_media_types.currentIndex())
        if media_type == 10:
            self.media_file_lookup_tvshow()
        elif media_type == 9:
            self.media_file_lookup_movie()
        else:
            QMessageBox.critical(self, "Error", f"No metadata lookup for {media_type_name} implemented.", QMessageBox.Ok)

    #
    # Save the tags.
    # 
    def save_file(self, file):
        indexes = self.media_file_view.selectedIndexes()
        for index in indexes:
            mediafile = self.model.mediafiles[index.row()]
            print (f"Save file {mediafile.file}")
            print (mediafile.metadata['tags']['title'])
            xml = mediafile.GenerateXML()
            _, tmp_file = tempfile.mkstemp()
            with open(tmp_file, 'w') as f:
                f.write(str(xml))
            try:
                result = subprocess.run(['mkvpropedit', '--gui-mode', str(mediafile.file), '--tags', 
                                        'global:' + str(tmp_file)], capture_output=True)
            except subprocess.CalledProcessError as err:
                print (f"ERROR: file tags not written {err}")
            if result.returncode != 0:
                print(result)
            else:
                print (f"File saved: {result.stdout}")
            try:
                reult =  subprocess.run(['mkvpropedit', '--gui-mode', str(mediafile.file), '--edit', 'info', '--set',
                                     f"title={mediafile.metadata['tags']['title']}"], capture_output=True)
            except subprocess.CalledProcessError as err:
                print (f"ERROR: title not set. {err}")
            if result.returncode != 0:
                print (result)
            else:
                print (f"Title set: {result.stdout}")
            mediafile.changes = False
            os.remove(tmp_file)
    #
    # Close the currently selected file
    #
    def close_file(self):
        indexes = self.media_file_view.selectedIndexes()
        for index in indexes:
            file  = self.model.mediafiles[index.row()]
            if file.changes:
                button = QMessageBox.warning(self, "Unsaved Tags!", f"{os.path.basename(file.file)} has unsaved changes to the tags save them?", 
                                          QMessageBox.Discard | QMessageBox.Save, defaultButton=QMessageBox.Discard)
                if button == QMessageBox.Discard:
                    del self.model.mediafiles[index.row()]
                else:
                    self.save_file(file)
                    del self.model.mediafiles[index.row()]
                self.media_file_view.clearSelection()
            else:
                del self.model.mediafiles[index.row()]
                self.media_file_view.clearSelection()

        # nothing selected.


    def open_files(self):
        QApplication.setOverrideCursor(Qt.WaitCursor)
        files, _ = QFileDialog.getOpenFileNames(self, 'Open Media files', self.current_path, 'Media Files (*.mkv *.mka)')
        for file in files:
            mediafile = MediaFile(file)
            mediafile.metadata['changed'] = False
            self.model.mediafiles.append(mediafile)
            self.model.layoutChanged.emit()

        # Select the last added item if no selection otherwise don't
        indexes = self.media_file_view.selectedIndexes()
        if indexes:
            # Something selected. leave it be.
            QApplication.restoreOverrideCursor()
            return
        # Select the last file added.
        file_count = self.model.rowCount(QModelIndex)
        index = self.model.index(file_count - 1,0)
        self.media_file_view.selectionModel().setCurrentIndex(index,QItemSelectionModel.SelectCurrent)
        QApplication.restoreOverrideCursor()


app = QtWidgets.QApplication(sys.argv)
window = MainWindow()
window.show()
app.exec_()

