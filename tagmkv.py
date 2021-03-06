#!/usr/bin/env python3
from __future__ import print_function, unicode_literals

from PyQt5.QtWidgets import *
from PyQt5 import QtCore, QtGui
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.uic import loadUi
from tmdbv3api import *
from enum import Enum

import lxml.etree as ET
import pprint
import sys
import argparse
import logging
import subprocess
import re
import os
import json
import unicodedata
import tempfile
import xmltodict
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

##
## Parse the filename and try and get a suitable title from it.
valid_extensions = (".avi", ".xvid", ".ogv", ".ogm", ".mkv", ".mpg", ".mp4", ".srt")
# cruft we want to strip
audio = ['([^0-9])5\.1[ ]*ch(.)','([^0-9])5\.1([^0-9]?)','([^0-9])7\.1[ ]*ch(.)','([^0-9])7\.1([^0-9])']
video = ['3g2', '3gp', 'asf', 'asx', 'avc', 'avi', 'avs', 'bivx', 'bup', 'divx', 'dv', 'dvr-ms', 'evo', 'fli', 'flv',
         'm2t', 'm2ts', 'm2v', 'm4v', 'mkv', 'mov', 'mp4', 'mpeg', 'mpg', 'mts', 'nsv', 'nuv', 'ogm', 'ogv', 'tp',
         'pva', 'qt', 'rm', 'rmvb', 'sdp', 'svq3', 'strm', 'ts', 'ty', 'vdr', 'viv', 'vob', 'vp3', 'wmv',
         'wtv', 'xsp', 'xvid', 'webm']
format= ['ac3','dc','divx','fragment','limited','ogg','ogm','ntsc','pal','ps3avchd','r1','r3','r5',
         'remux','x264','xvid','vorbis','aac','dts','fs','ws','1920x1080',
         '1280x720','h264','h','264','prores','uhd','2160p','truehd','atmos','hevc']              
misc  = ['cd1','cd2','1cd','2cd','custom','internal','repack','read.nfo','readnfo','nfofix','proper',
         'rerip','dubbed','subbed','extended','unrated','xxx','nfo','dvxa']
subs  = ['multi', 'multisubs']
sizes = ['480p', '720p', '1080p', '480i', '720i', '1080i']
src_dict = {'bluray':['bdrc','bdrip','bluray','bd','brrip','hdrip','hddvd','hddvdrip'],
                        'cam':['cam'],'dvd':['ddc','dvdrip','dvd','r1','r3','r5'],'retail':['retail'],
                        'dtv':['dsr','dsrip','hdtv','pdtv','ppv'],'stv':['stv','tvrip'],
                        'screener':['bdscr','dvdscr','dvdscreener','scr','screener'],
                        'svcd':['svcd'],'vcd':['vcd'],'telecine':['tc','telecine'],
                        'telesync':['ts','telesync'],'web':['webrip','web-dl'],'workprint':['wp','workprint']}
yearRx = '([\(\[ \.\-])([1-2][0-9]{3})([\.\-\)\]_,+])'
sizeRx = '([0-9]{3,4}[i|p])'

source = []
for d in src_dict:
    for s in src_dict[d]:
        if source != '':
            source.append(s);
reversed_tokens = set()
for f in format + source:
    if len(f) > 3:
        reversed_tokens.add(f[::-1].lower())


def titlecase(s):
    return re.sub(
        r"[A-Za-z]+('[A-Za-z]+)?",
        lambda word: word.group(0).capitalize(),
        s)

def GetYear(name):
    # Grab the year if specified
    year = None
    yearMatch = re.search(yearRx, name)
    if yearMatch:
        yearStr = yearMatch.group(2)
        yearInt = int(yearStr)
        if yearInt > 1900 and yearInt < (datetime.date.today().year + 1):
            year = int(yearStr)
            name = name.replace(yearMatch.group(1) + yearStr + yearMatch.group(3), ' *yearBreak* ')
    return Year

def CleanName(name):
    name_tokens_lowercase = set()
    for t in re.split('([^ \-_\.\(\)+]+)', name):
        t = t.strip()
        if not re.match('[\.\-_\(\)+]+', t) and len(t) > 0:
            name_tokens_lowercase.add(t.lower())
    if len(set.intersection(name_tokens_lowercase, reversed_tokens)) > 2:
        name = name[::-1]
    orig = name
    try:
        name = unicodedata.normalize('NFKC', name.decode(sys.getfilesystemencoding()))
    except:
        try:
            name = unicodedata.normalize('NFKC', name.decode('utf-8'))
        except:
            pass
    name = name.lower()
    
    # Grab the size if it's one that we want
    size = None
    sizeMatch = re.search(sizeRx, name)
    if sizeMatch:
        size = sizeMatch.group(1)
        print ("Resolution detected as %s" % (size))
        
    # Take out things in brackets. (sub acts weird here, so we have to do it a few times)
    done = False
    while done == False:
        (name, count) = re.subn(r'\[[^\]]+\]', '', name, re.IGNORECASE)
        if count == 0:
            done = True
    # Take out audio specs, after suffixing with space to simplify rx.
    name = name + ' '
    for s in audio:
        rx = re.compile(s, re.IGNORECASE)
        name = rx.sub(' ', name)

    # Now we tokenize it
    tokens = re.split('([^ \-_\.\(\)+]+)', name)

    # Process the tokens
    newTokens = []
    tokenBitmap = []
    seenTokens = {}
    finalTokens = []

    for t in tokens:
            t = t.strip()
            if not re.match('[\.\-_\(\)+]+', t) and len(t) > 0:
                newTokens.append(t)
    if newTokens[-1] != '*yearBreak*':
        extension = "." + newTokens[-1]
    else:
        extension = None
    # Now we build a bitmap of what we want to keep (good) and what to toss (bad)
    garbage = subs
    garbage.extend(misc)
    garbage.extend(format)
    garbage.extend(source)
    garbage.extend(video)
    garbage = set(garbage)
    for t in  reversed(newTokens):
        if t.lower() in garbage and t.lower() not in seenTokens:
            seenTokens[t.lower()] = True
            tokenBitmap.insert(0, False)
        else:
            tokenBitmap.insert(0,True)
    numGood = 0
    numBad  = 0
    for i in range(len(tokenBitmap)):
        good = tokenBitmap[i]
        if len(tokenBitmap) <= 2:
            good = True
        if good and numBad <= 2:
            if newTokens[i] =='*yearBreak*':
                if i == 0:
                    continue
                else:
                    break
            else:
                finalTokens.append(newTokens[i])
        elif not good and newTokens[i].lower() == 'dc':
            if i+1 < len(newTokens) and newTokens[i+1].lower() in ['comic', 'comics']:
                finalTokens.append('DC')
            else:
                finalTokens.append("(Director's cut)")
    if good == True:
        numGood += 1
    else:
        numBad += 1
    if len(finalTokens) == 0 and len(newTokens) > 0:
        finalTokens.append(newTokens[0])

    cleanedName = ' '.join(finalTokens)
    
#    if extension and extension in valid_extensions:
#        if size:
#            cleanedName += " - " + size
#        cleanedName += extension
    return (titlecase(cleanedName))

def lowercase_keys(obj):
  if isinstance(obj, dict):
    obj = {key.lower(): value for key, value in obj.items()}
    for key, value in obj.items():         
      if isinstance(value, list):
        for idx, item in enumerate(value):
          value[idx] = lowercase_keys(item)
      obj[key] = lowercase_keys(value)
  return obj 

### Class Property
####################################################
class MediaProperty():
    def __init__(self, name, value):
        self.name = name
        self.value = value
        self.child = None

    def setChild(self, name, value):
        self.child = Property(name, value)
        return 0

    def getChild(self):
        return self.child


### Class MediaFile
####################################################
class MediaFile():
    def __init__(self, filename):
        ## Existing properties on the file
        self.properties = []
        self.filename = filename
        self.dirname = ''
        self.fullname = ''
        self.metadata = None
        self.img = ''

    def __eq__(self, filename):
        return self.filename == filename

### Search Results dialog.
class SearchResults(QDialog):
    def __init__(self, results):
        super(SearchResults, self).__init__()
        self.w = loadUi('ui/tv_lookup_results.ui', self)
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
        self.w.show()

    def getSelectedResult(self):
        selected = self.ResultsList.item(self.ResultsList.currentRow())
        return selected

    def ResultsListClicked(self, item):
        result = item.data(Qt.UserRole)
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

###
### Main Window
##########################################################################
class Window(QMainWindow):

    ## Stuff to handle extracting the tag info out of the XML format mkv uses.
    tag_xpath = '/Tags/Tag/Simple/Name[string() = $name]/..'
    metadata_tags = ("SHOW", "SUMMARY", "SEASON", "EPISODE", 
                     "TITLE", "DESCRIPTION", "TMDB", "MEDIA_TYPE", "GENRE", "ACTOR", "media_type")

    def __init__(self, parent=None):
        super().__init__(parent)
        loadUi('ui/main_window.ui', self)
        logging.basicConfig(encoding='utf-8', level=logging.INFO)

        self.logger = logging.getLogger("MainWindow")
        self.logger.setLevel(logging.INFO)

        self.cast_model = QStandardItemModel()
        self.cast_model.setHorizontalHeaderLabels(['Actor', 'Character'])
        self.CastView.setModel(self.cast_model)
        self.CastView.horizontalHeader().setSectionResizeMode(1)
        self.media_files = []
        self.current_file = ''
        self.current_path = os.environ['HOME']
        self.tmdb = TMDb()
        self.tmdb.language = 'en'
        config = Configuration()
        self.tmdb_config = config.info()
        self.tv_genres = self.get_tv_genres()
        self.movie_genres = self.get_movie_genres()
        self.setup_media_types()
        self.setupUi()
        self.connect_signals()

    def setupUi(self):
        self.TVShowGroup.setEnabled(False)
        self.MediaType.setCurrentIndex(self.MediaType.findData(9))
        return

    def setLogLevel(self, level):
        numeric_level = getattr(logging, level.upper(), None)
        if not isinstance(numeric_level, int):
            raise ValueError('Invalid log level: %s' % level)
        self.logger.setLevel(numeric_level)
        return 

    def connect_signals(self):
        ### File Menu
        self.actionOpenFile.triggered.connect(self.OpenFile)
        self.actionCloseFile.triggered.connect(self.CloseFile)
        self.actionSaveFile.triggered.connect(self.SaveFile)
        self.actionExit.triggered.connect(self.close)
        ### Interface Widgets
        self.FileList.currentItemChanged.connect(self.FileListChanged)
        ##
        self.MediaType.activated.connect(self.MediaTypeActivated)
        self.TVShow.editingFinished.connect(self.TVShowEditingFinished)
        self.TVSeason.valueChanged.connect(self.TVSeasonChanged)
        self.TVEpisode.valueChanged.connect(self.TVEpisodeChanged)
        self.MetadataLookup.clicked.connect(self.MetadataLookupClicked)
        ##
        self.MediaTitle.editingFinished.connect(self.MediaTitleEditingFinished)
        self.Genres.clicked.connect(self.GenreListClicked)

    def setup_media_types(self):
        for key, value in media_types.items():
            self.MediaType.addItem(value,key)

    def SetMediaType(self, media_type):
        self.logger.debug("SetMediaType")
        mediafile = self.getMediaFile()
        tags = mediafile.metadata['format']['tags']
        tags['media_type'] = int(self.MediaType.currentData())


    def title_from_filename(self, mediafile):
        ## set the metadata title tag from the filename 
        self.logger.debug("title_from_filename")
        ftitle = CleanName(mediafile.filename)
        tags = mediafile.metadata['format']['tags']
        tags['title'] = ftitle

    ### Simplistic attempt to get the name, season, episode and episode_title from the filename
    ### it expects things in the following format:
    ### <show> S<season>E<episode> - <title>.[xxx]
    ### <show> <season>x<episode> <title>.[xxx]
    ### (?P<series>.*)\s[Ss]?(?P<season>\d+)[Ee|\s|x]?(?P<episode>\d+)\W+(?P<title>.*)\.\w{3}$
    def tv_show_from_filename(self, mediafile):
        self.logger.debug("tv_show_from_filename (%s)", mediafile.filename)
        self.StatusBar.showMessage("tv_show_from_filename (%s)" % mediafile.filename)
        pattern = re.compile('(?P<series>.*)\s[Ss]?(?P<season>\d+)[Ee|\s|x]?(?P<episode>\d+)\W+(?P<title>.*)\.\w{3}$')
        result = pattern.search(mediafile.filename)
        if result:
            self.logger.debug("tv_show_from_filename (full match) %s", result.groupdict())
            return result.groupdict()
        pat = re.compile('(?P<series>.*)\s[Ss]?(?P<season>\d+)[Ee|x]?(?P<episode>\d+)\.\w{3}')
        result = pat.search(mediafile.filename)
        if result:
            self.logger.debug("tv_show_from_filename (no title) %s", result.groupdict())
            return result.groupdict()
        else:
            return None 

    def get_tv_genres(self):
        self.logger.debug("get_tv_genres")
        tmdb = Genre()
        return tmdb.tv_list()

    def get_movie_genres(self):
        self.logger.debug('get_movie_genres')
        tmdb = Genre()
        return tmdb.movie_list()

    def UpdateGenre(self, genre_tag, media_type):
        self.logger.debug("UpdateGenre %s", genre_tag)
        self.Genres.clear()
        self.GenreTag.clear()
        self.GenreTag.setText(genre_tag)
        if media_type == '10':
            for tv_genre in self.tv_genres:
                self.logger.debug("TV Genre %s", tv_genre)
                item = QListWidgetItem(tv_genre['name'])
                self.Genres.addItem(item)
        else:
            for movie_genre in self.movie_genres:
                self.logger.debug("Movie Genre %s", movie_genre)
                item = QListWidgetItem(movie_genre['name'])
                self.Genres.addItem(item)
        genres = self.unpack_media_genres(genre_tag)
        for genre in genres:
            self.logger.debug("Genre: %s", genre)
            matches = self.Genres.findItems(genre, Qt.MatchContains)
            for match in matches:
                match.setSelected(True)


    def pack_media_genres(self, tags):
        self.logger.debug("pack_media_genres")
        genres = list()
        for genre in tags:
            self.logger.debug("Genre %s", genre['name'])
            genres.append(genre['name'])
        genre_tag = '|'.join(genres)
        self.logger.debug("genre_tag %s", genre_tag)
        return (genre_tag)

    def unpack_media_genres(self, genre_tag):
        self.logger.debug("unpack_media_genres %s", genre_tag)
        genres = genre_tag.split('|')
        return genres

    ###
    ### TMDB functions
    ######################################################################
    #
    ### https://developers.themoviedb.org/3/search/search-movies
    def FindMovieMetadata(self, mediafile):
        self.logger.debug("Search for movie metadata (%s)", mediafile.filename)
        tmdb = Movie()
        tags = mediafile.metadata['format']['tags']
        try:
            results = tmdb.search(tags['title'])
        except TMDBException:
            self.StatusBar.showMessage("Movie '%s' not found", tags['title'])
            return
        if len(results) == 1:
            ## One result, must be what we were looking for.
            self.logger.debug("One result %s", results[0]['id'])
            tmdb_id = results[0]['id']
            self.GetMovieMetadata(tmdb_id)
        else:
            self.logger.debug("Got %s results, open selection dialog", len(results))
            self.resultsDialog = SearchResults(results)
            self.resultsDialog.buttonBox.accepted.connect(self.SelectedMovieMetadata)

    #### https://developers.themoviedb.org/3/movies/get-movie-details  
    def GetMovieMetadata(self, tmdb_id):
        self.logger.debug("GetMovieMetadata (%s)", tmdb_id)
        mediafile = self.getMediaFile()
        tags = mediafile.metadata['format']['tags']

        tmdb = Movie()
        movie = tmdb.details(tmdb_id, append_to_response='credits')
        ## Set the tags
        tags['title'] = movie['title']
        tags['tmdb'] = 'movie/' + str(tmdb_id)
        tags['description'] = movie['overview']
        tags['date_released'] = movie['release_date']
        tags['genre'] = self.pack_media_genres(movie['genres'])
        credits = movie['credits']
        tmdb_cast = credits['cast']
        tags_cast = tags['cast']
        tags_cast.clear()
        self.cast_model.removeRows(0, self.cast_model.rowCount())
        for cast_member in tmdb_cast:
            row = (QStandardItem(cast_member['name']), QStandardItem(cast_member['character']))
            tags_cast.append({'actor': cast_member['name'], 'character': cast_member['character']})
            self.cast_model.appendRow(row)

        ## Refresh the UI
        self.TMDBID.setValue(tmdb_id)
        self.MediaTitle.setText(tags['title'])
        self.MediaDescription.setPlainText(tags['description'])
        d = QDate.fromString(tags['date_released'], 'yyyy-MM-dd')
        self.ReleasedDate.setDate(d)
        self.UpdateGenre(tags['genre'], tags['media_type'])

    def SelectedMovieMetadata(self):
        self.logger.debug("SelectedMovieMetadata")
        mediafile = self.getMediaFile()
        item = self.resultsDialog.getSelectedResult()
        selected_movie = item.text()
        tmdb_id = item.data(Qt.UserRole)['id']
        self.GetMovieMetadata(tmdb_id)

        # movie = Movie()
        # details = movie.details(tmdb_id)
        # ### Set tags
        # tags = mediafile.metadata['format']['tags']
        # tags['tmdb'] = 'movie/' + str(tmdb_id)
        # tags['description'] = details['overview']
        # tags['date_released'] = details['release_date']
        # if 'genres' in details:
        #     tags['genre'] = self.pack_media_genres(details['genres'])
        # ### Update UI
        # self.MediaDescription.setPlainText(details['overview'])
        # d = QDate.fromString(tags['date_released'], 'yyyy-MM-dd')
        # self.ReleasedDate.setDate(d)

    ### https://developers.themoviedb.org/3/tv-seasons/get-tv-season-details
    def getShowEpisode(self, tmdb_id):
        self.logger.debug("getShowEpisode %d", tmdb_id)
        mediafile = self.getMediaFile()
        tags = mediafile.metadata['format']['tags']
        ### Make sure we are marked as a TV Show media type.
        tags['media_type'] = 10
        ## Show details
        tv = TV()
        show = tv.details(tmdb_id)
        ## Gepisode Season and Number
        tv_season  = int(self.TVSeason.value())
        tv_episode = int(self.TVEpisode.value())
        ## Get the details
        episode = Episode()
        episode_details = episode.details(tmdb_id, tv_season, tv_episode, append_to_response="credits")
        ### Set tags
        tags['show'] = show['name']
        tags['summary'] = show['overview']
        tags['tmdb'] = 'tv/' + str(self.TMDBID.value())
        tags['title'] = episode_details['name']
        tags['description'] = episode_details['overview']
        tags['date_released'] = episode_details['air_date']
        if 'genres' in show:
            tags['genre'] = self.pack_media_genres(show['genres'])
        credits = episode_details['credits']
        tmdb_cast = credits['cast']
        tags_cast = tags['cast']
        tags_cast.clear()
        self.cast_model.removeRows(0, self.cast_model.rowCount())
        for cast_member in tmdb_cast:
            row = (QStandardItem(cast_member['name']), QStandardItem(cast_member['character']))
            tags_cast.append({'actor': cast_member['name'], 'character': cast_member['character']})
            self.cast_model.appendRow(row)
        ### Update UI
        self.TVShowSummary.setText(tags['summary'])
        self.MediaDescription.setPlainText(tags['description'])
        self.MediaTitle.setText(tags['title'])
        self.TMDBID.setValue(tmdb_id)
        d = QDate.fromString(tags['date_released'], 'yyyy-MM-dd')
        self.ReleasedDate.setDate(d)
        self.UpdateGenre(tags['genre'], tags['media_type'])

    ### https://developers.themoviedb.org/3/search/search-tv-shows
    def FindTVMetadata(self, mediafile):
        self.logger.debug("FindTVMetadata (%s)", mediafile.filename)
        self.StatusBar.showMessage("Lookup tv metadata for {}".format(mediafile.filename))
        tags = mediafile.metadata['format']['tags']
        ### We have an ID, go ahead and fetch that ID
        if 'tmdb' in tags:
            _,tmdb_id = tags['tmdb'].split('/')
            self.logger.debug("Previously set tmdb_id {}, get those details.".format(tmdb_id))
            self.getShowEpisode(int(tmdb_id))
        else:
            self.logger.debug("FindTVMetaData() search for {}".format(tags['show']))
            tv = TV()
            show = tv.search(tags['show'])
            if len(show) == 1:
                tmdb_id = show[0]['id']
                self.getShowEpisode(tmdb_id)
            else:
                self.resultsDialog = SearchResults(show)
                self.resultsDialog.buttonBox.accepted.connect(self.SelectedTVMetadata)

    def SelectedTVMetadata(self):
        self.logger.debug("SelectedTVMetadata")
        item = self.resultsDialog.getSelectedResult()
        selected_show = item.text()
        tmdb_id = item.data(Qt.UserRole)['id']
        self.getShowEpisode(tmdb_id)

    ####
    # ffmpeg can handle 99% of this
    # ffprobe -v error -of default=noprint_wrappers=0 -print_format json -show_format -show_streams -show_chapters -i <file>
    # 
    # mvkextract can also pull this out but in xml.
    # mkvextract <file> tags --global-tags tags.xml
    #

    def AnalyzeFile(self, mediafile):
        self.StatusBar.showMessage("Analyzing file {}".format(mediafile.filename))
        self.logger.debug("AnalyzeFile (%s)", mediafile.filename)
        try:
            output = subprocess.run(['ffprobe', mediafile.fullname,
                                     '-v', 'error', 
                                     '-hide_banner', 
                                     '-of', 'default=noprint_wrappers=0', 
                                     '-print_format', 'json', 
                                     '-show_format', '-show_streams', '-show_chapters' 
                                     ], text=True, check=True, capture_output=True, universal_newlines=True)
        except subprocess.CalledProcessError as Err:
            self.StatusBar.showMessage('ERROR: File is unreadable')
            return 1
        metadata = json.loads(output.stdout)

        _, temp_file_path = tempfile.mkstemp(suffix='.xml')
        try:
            output = subprocess.run(['mkvextract', mediafile.fullname,
                                    'tags',
                                    '--global-tags',
                                    temp_file_path], text=True, check=True, capture_output=True, universal_newlines=True)
        except subprocess.CalledProcessError as Err:
            self.StatusBar.showMessage('ERROR: File is unreadable')
            return 1
        root = ET.parse(temp_file_path).getroot()
        os.remove(temp_file_path)
        ## Build the tags from the XML data, with only the tags we are interested in to replace the broken ffprobe
        ## parsing of nested tags.
        ## ffprobe (avcodec?) _does not_ handle:
        ##  Actor Name
        ##      Character
        ##  Actor Name
        ##      Character
        ## ...
        ##     
        xml_tags = dict()
        xml_tags['cast'] = []
        actors = xml_tags['cast']
        for tag in self.metadata_tags:
            elem = root.xpath(self.tag_xpath, name=tag)
            if elem:
                for item in elem:
                    for sub_elem in item.getchildren():
                        if sub_elem.text == 'ACTOR':
                            parent = item.xpath('Simple/String')
                            value = item.xpath('String')
                            actor = {str(tag): value[0].text }
                            if parent:
                                if parent[0] is not None:
                                    actor['CHARACTER'] = parent[0].text
                            xml_tags['cast'].append(actor)
                            break
                        if sub_elem.tag == 'String':
                            xml_tags[str(tag)] = sub_elem.text
        metadata['format']['tags'] = lowercase_keys(xml_tags)

        ## Matroska metadata puts the keys in UPPERCASE while other formats have them in lower case
        ## lets standardize on lowercase.
        #metadata['format']['tags'] = {k.casefold(): v for k, v in metadata['format']['tags'].items()}
        return metadata

    def ResetTVShow(self):
        self.logger.debug("ResetTVShow")
        tags = self.getMediaFile().metadata['format']['tags']
        self.TVShow.clear()
        ## Special handling of the season/episode spinboxes
        if 'season' not in tags:
            self.TVSeason.setValue(0)
        if 'episode' not in tags:
            self.TVEpisode.setValue(0)
        self.TVShowSummary.clear()
        self.cast_model.removeRows(0, self.cast_model.rowCount())

    def UpdateTVShow(self):
        mediafile = self.getMediaFile()
        tags = mediafile.metadata['format']['tags']
        self.logger.debug("UpdateTVShow (%s)", mediafile.filename)
        if 'summary' in tags:
            self.TVShowSummary.setPlainText(tags['summary'])
        if 'show' in tags:
            self.TVShow.setText(tags['show'])
        else:
            self.TVShow.clear()
        if 'season' in tags:
            self.TVSeason.setValue(int(tags['season']))
        else:
            self.TVSeason.setValue(0)
        if 'episode' in tags:
            self.TVEpisode.setValue(int(tags['episode']))
        else:
            self.TVEpisode.setValue(0)

    def ProcessFile(self, mediafile):
        self.StatusBar.showMessage("Load previously analyzed file {}".format(mediafile.filename))
        self.logger.debug("ProcessFile (%s)", mediafile.filename)
        if 'tags' in mediafile.metadata['format']:
            tags = mediafile.metadata['format']['tags']
        else:
            ### This should never happen as ffmpeg will populate the 'tags' dict with at the very least
            ### the encoder tag
            tags = dict()
            mediafile.metadata['format']['tags'] = tags
        ### No media_type default to a 'movie' media_type
        media_type = 9
        if 'media_type' not in tags:
            tags['media_type'] = 9
        else:
            media_type = int(tags['media_type'])
        self.MediaType.setCurrentIndex(self.MediaType.findData(int(tags['media_type'])))
        if media_type == 10:
            self.logger.debug("media_type = TVshow")
            ### Set TV Show specific tags
            self.TVShowGroup.setEnabled(True)
            self.logger.debug("UpdateUI")
            self.ResetTVShow()
            self.UpdateTVShow()
        else:
            self.logger.debug("media_type = other")
            self.ResetTVShow()
            self.TVShowGroup.setEnabled(False)

        ### Set release date
        if 'date_released' in tags:
            d = QDate.fromString(tags['date_released'], 'yyyy-MM-dd')
            self.ReleasedDate.setDate(d)
        else:
            d = QDate.currentDate()
            self.ReleasedDate.setDate(d)
            tags['date_released'] = d.toString(Qt.ISODate)

        ## common metadata for movie/tv show
        if 'title' not in tags:
            tags['title'] = self.title_from_filename(mediafile)
        self.MediaTitle.setText(tags['title'])
        # description, also known as synopsis
        if 'description' in tags:
            self.MediaDescription.setPlainText(tags['description'])
        else:
            self.MediaDescription.clear()
        self.GenreTag.clear()
        self.Genres.clear()
        if 'genre' in tags:
            self.UpdateGenre(tags['genre'], tags['media_type'])

        if 'tmdb' in tags:
            prefix, tmdb_id = tags['tmdb'].split('/')
            self.TMDBID.setValue(int(tmdb_id))
        else:
            self.TMDBID.clear()
        if 'cast' in tags:
            for member in tags['cast']:
                if 'actor' in member and 'character' in member:
                    row = (QStandardItem(member['actor']), QStandardItem(member['character']))
                    self.cast_model.appendRow(row)
                elif 'actor' in member:
                    row = (QStandardItem(member['actor']))
                    self.cast_model.appendRow(row)               

    def addMediaFile(self, filename):
        self.logger.debug("addMediaFile (%s)", filename)
        fname = os.path.basename(filename)
        for mediafile in self.media_files:
            if mediafile.filename == fname:
                self.logger.debug("Existing mediafile (%s), not adding", filename)
                return mediafile
        else:
            self.logger.debug("New mediafile, append to list (%s)", filename)
            mediafile = MediaFile(os.path.basename(filename))
            mediafile.dirname = os.path.dirname(filename)
            mediafile.fullname = filename
            mediafile.metadata = self.AnalyzeFile(mediafile)
            self.media_files.append(mediafile)
        return mediafile

    def getMediaFile(self):
        self.logger.debug("getMediaFile")
        return self.FileList.currentItem().data(Qt.UserRole)

    ###
    ### Interface actions
    ###
    def GenreListClicked(self, index):
        item = self.Genres.itemFromIndex(index)
        self.logger.debug("GenreListClicked %s", item.text())
        selected_items = self.Genres.selectedItems()
        genres = list()
        # Build new genre dict 
        for selected_item in selected_items:
            genres.append({'name': selected_item.text()})
        genre_tag = self.pack_media_genres(genres)
        self.logger.debug("New genre_tag %s", genre_tag)
        ## Set the new genre_tag
        mediafile = self.getMediaFile()
        tags = mediafile.metadata['format']['tags']
        tags['genre'] = genre_tag
        ## Update the UI
        self.UpdateGenre(genre_tag, tags['media_type'])

    def FileListChanged(self, item):
        mediafile = item.data(Qt.UserRole)
        self.logger.debug("FileListChanged (%s)", mediafile.filename)
        mediafile = item.data(Qt.UserRole)
        self.current_file = mediafile.filename
        self.current_path = mediafile.dirname
        self.CurrentDirectory.setText(self.current_path)
        self.CurrentFile.setText(self.current_file)
        self.ProcessFile(mediafile)

    def MediaTypeActivated(self, index):
        self.logger.debug("MediaTypeActivated (%d)", index)
        media_type = int(self.MediaType.itemData(index))
        mediafile = self.getMediaFile()
        tags = mediafile.metadata['format']['tags']
        if media_type == 10:
            tags['media_type'] = 10
            ## Attempt to autodetect info from the filename and set tags appropriately
            tv_show = self.tv_show_from_filename(mediafile)
            if 'series' in tv_show:
                tags['show'] = tv_show['series'] 
            if 'season' in tv_show:
                tags['season'] = tv_show['season']
            if 'episode' in tv_show:
                tags['episode'] = tv_show['episode']
            if 'title' in tv_show: 
                tags['title'] = tv_show['title']
            pprint.pprint(tags)
            self.logger.debug("UpdateUI")
            self.ResetTVShow()
            self.UpdateTVShow()
            self.TVShowGroup.setEnabled(True)
        else:
            self.ResetTVShow()
            self.TVShowGroup.setEnabled(False)
            tags['media_type'] = 9
        self.SetMediaType(mediafile)
        return 

    def TVShowEditingFinished(self):
        tvshow = self.TVShow.text()
        mediafile = self.getMediaFile()
        tags = mediafile.metadata['format']['tags']
        tags['show'] = tvshow
        return 

    def TVSeasonChanged(self):
        tv_season = self.TVSeason.value()
        mediafile = self.getMediaFile()
        tags = mediafile.metadata['format']['tags']
        tags['season'] = tv_season
        return

    def TVEpisodeChanged(self):
        tv_episode = self.TVEpisode.value()
        mediafile = self.getMediaFile()
        tags = mediafile.metadata['format']['tags']
        tags['episode'] = tv_episode
        return 

    def MetadataLookupClicked(self):
        mediafile = self.getMediaFile()
        media_type = int(self.MediaType.itemData(self.MediaType.currentIndex()))
        self.logger.info("media_type: {}".format(media_type))
        if media_type == 10:
            self.FindTVMetadata(mediafile)
        else:
            self.FindMovieMetadata(mediafile)

    def MediaTitleEditingFinished(self):
        title = self.MediaTitle.text()
        mediafile = self.getMediaFile()
        tags = mediafile.metadata['format']['tags']
        tags['title'] = title
        return 

    ### File Menu actions
    def OpenFile(self):
        self.logger.debug("OpenFile")
        QApplication.setOverrideCursor(Qt.WaitCursor)
        selected_row = 0
        if len(self.media_files) == 0:
            selected_row = -1 
        self.files, _ = QFileDialog.getOpenFileNames(self, 'Open Media files', self.current_path, 'Media Files (*.mkv *.mp4 *.m4a *.avi *.ogg)')
        for file in self.files:
            self.logger.debug("SelectedFile: (%s)", os.path.basename(file))
            mediafile = self.addMediaFile(file)
            item = QListWidgetItem(os.path.basename(file))
            item.setData(Qt.UserRole, mediafile)
            self.FileList.addItem(item)
        if selected_row < 0:
            selected_row = 0
        else:
            selected_row = self.FileList.count() - 1
        self.FileList.setCurrentRow(selected_row)
        QApplication.restoreOverrideCursor()
        return 0

    def SaveFile(self):
        mediafile = self.getMediaFile()
        xml_file = self.CreateXML(mediafile)
        if xml_file:
            try:
                output = subprocess.run(['mkvpropedit', '--gui-mode', str(mediafile.fullname), '--tags', 
                                        'global:' + str(xml_file)], capture_output=True)
            except subprocess.CalledProcessError as err:
                self.logger.error('Error: %s', err)
                self.StatusBar.showMessage("ERROR: MKV File tag's not written '{}'".format(str(err)))
            if (os.path.isfile(xml_file)):
                self.logger.debug("Delete temp file %s", xml_file)
                os.remove(xml_file)
        return 0

    def CloseFile(self):
        mediafile = self.getMediaFile()
        

    ## Create return the elements for a 'simple' tag 
    ## Which consists of 
    '''
        <Simple>
           <Name>{name}</Name>
           <String>{value}</String>
        </Simple>
        The TagName SHOULD always be written in all capital letters and contain no space.

        The fields with dates SHOULD have the following format: YYYY-MM-DD hh:mm:ss.mss 
        YYYY = Year, MM = Month, DD = Days, HH = Hours, mm = Minutes, ss = Seconds, mss = Milliseconds. 
        To store less accuracy, you remove items starting from the right. 
        To store only the year, you would use, ???2004???. 
        To store a specific day such as May 1st, 2003, you would use ???2003-05-01???.

        Fields that require a Float SHOULD use the ???.??? mark instead of the ???,??? mark. 
        To display it differently for another local, applications SHOULD support auto replacement on display. 
        Also, a thousandths separator SHOULD NOT be used.
        For currency amounts, there SHOULD only be a numeric value in the Tag. 
        Only numbers, no letters or symbols other than ???.???. For instance, you would store ???15.59??? instead of ???$15.59USD???.
        https://www.matroska.org/technical/tagging.html
    '''
    def SimpleTag(self, tag):
        self.logger.debug("Encode %s", tag)
        for k,v in tag.items():
            simple = ET.Element('Simple')
            name = ET.Element('Name')
            name.text = str(k).upper()
            simple.append(name)
            string = ET.Element('String')
            string.text = str(v)
            simple.append(string)
        return simple


    def CreateXML(self, mediafile):
        self.logger.debug("CreateXML tag file")
        metadata = mediafile.metadata['format']['tags']
        root = ET.Element('Tags')
        tag  = ET.Element('Tag')
        root.append(tag)
        targets = ET.Element('Targets')
        tag.append(targets)
        ## Set the media_type
        if 'media_type' not in metadata:
            tags['media_type'] = 9
        targettypevalue = ET.Element('TargetTypeValue')
        targettypevalue.text = '70'
        targets.append(targettypevalue)
        simple = self.SimpleTag({'media_type': metadata['media_type']})
        tag.append(simple)
        ## Set Show information
        if 'show' in metadata:
            tag.append(self.SimpleTag({'show': metadata['show']}))
        if 'summary' in metadata:
            tag.append(self.SimpleTag({'summary': metadata['summary']}))
        # Season tags
        if 'season' in metadata:
            tag = ET.Element('Tag')
            targets = ET.Element('Targets')
            root.append(tag)
            tag.append(targets)
            targettypevalue = ET.Element('TargetTypeValue')
            targettypevalue.text = '60'
            targets.append(targettypevalue)
            tag.append(self.SimpleTag({'season': metadata['season']}))

        # Now we have the episode/movie detail
        tag = ET.Element('Tag')
        targets = ET.Element('Targets')
        root.append(tag)
        tag.append(targets)
        targettypevalue = ET.Element('TargetTypeValue')
        targettypevalue.text = '50'
        targets.append(targettypevalue)
        if 'episode' in metadata:
            tag.append(self.SimpleTag({'episode': metadata['episode']}))
        if 'title' in metadata:
            tag.append(self.SimpleTag({'title': metadata['title']}))
        if 'description' in metadata:
            tag.append(self.SimpleTag({'description': metadata['description']}))
        if 'date_released' in metadata:
            tag.append(self.SimpleTag({'date_released': metadata['date_released']}))
        if 'genre' in metadata:
            tag.append(self.SimpleTag({'genre': metadata['genre']}))
        if 'tmdb' in metadata:
            tag.append(self.SimpleTag({'TMDB': metadata['tmdb']}))
        ## Add cast members
        for actor in metadata['cast']:
            if 'actor' in actor:
                cast_member = self.SimpleTag({'actor': actor['actor']})
            if 'character' in actor:
                cast_member.append(self.SimpleTag({'character': actor['character']}))
                tag.append(cast_member)

        _, temp_file_path = tempfile.mkstemp()
        # print(ET.tostring(root, xml_declaration=True, encoding = 'utf-8', pretty_print=True,
        #                   doctype = '<!DOCTYPE Tags SYSTEM "matroskatags.dtd">').decode('utf-8'))
        with open(temp_file_path, 'w') as file:
            file.write(ET.tostring(root, xml_declaration=True, encoding = 'utf-8', pretty_print=True,
                              doctype='<!DOCTYPE Tags SYSTEM "matroskatags.dtd">').decode('utf-8'))
        return temp_file_path

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Tag Media files with metadata from the Internet.')
    parser.add_argument('--log', type=str, dest='loglevel', default="INFO")
    args = parser.parse_args()
    app = QApplication(sys.argv)
    win = Window()
    win.setLogLevel(args.loglevel)
    win.show()
    sys.exit(app.exec())
