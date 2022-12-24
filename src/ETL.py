import spotipy as SP
from spotipy.oauth2 import SpotifyOAuth as OAuth
import spotipy.util as util
import pandas as pd
import secret
import sys


##Extracting Spotify Data
spotify = SP.Spotify(auth_manager=OAuth(client_id=secret.user_ID, client_secret=secret.user_sec,redirect_uri='http://localhost:9999', scope='user-top-read'))

topTracks = spotify.current_user_top_tracks(limit=50, offset=0, time_range='medium_term')

if len(topTracks) == 0:
    sys.exit("No tracks recieved from spotify, now exiting.")

counter = 1
song_info = {}
for track in topTracks['items']:
    

    artist_uri = track["artists"][0]["uri"]
    artist_info = spotify.artist(artist_uri)
    
    if counter == 40:
        song_info.update({counter:{'name': track['name'], 'artist':track['artists'][0]['name'], 'genre':"ccm"}})
        #print(str(counter) + ".) " + track['name'] + ", Artist: " + track['artists'][0]['name'] + ", Genre: ccm")
        counter += 1
    else:
        song_info.update({counter:{'name': track['name'], 'artist':track['artists'][0]['name'], 'genre':artist_info['genres'][0]}})
        #print(str(counter) + ".) " + track['name'] + ", Artist: " + track['artists'][0]['name'] + ", Genre: " + artist_info['genres'][0])
        counter += 1

##Transforming the Data

counter = 0
name_dic = {}
artist_dic = {}
genre_dic = {}

for key, songs in song_info.items():
    for i in songs:
        if str(i) == 'name':
            name_dic[str(int((counter/3)+1))] = songs[i]
        counter += 1

counter = 0
for key, songs in song_info.items():
    for i in songs:
        if str(i) == 'artist':
            artist_dic[str(int((counter/3)+1))] = songs[i]
        counter += 1

counter = 0
for key, songs in song_info.items():
    for i in songs:
        if str(i) == 'genre':
            genre_dic[str(int((counter/3)+1))] = songs[i]
        counter += 1

data = pd.DataFrame({})
data['name'] = name_dic 
data['artist'] = artist_dic
data['genre'] = genre_dic
print(data)

##Loading the Data to an excel sheet to perform more analysis on the data

file_name = 'Spotify_Data.xlsx'
data.to_excel(file_name)
print("Data written to file successfully")