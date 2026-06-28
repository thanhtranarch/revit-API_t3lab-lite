import os
import sys
import json
import time
import re

def check_dependencies():
    """Checks and prompts to install dependencies if missing."""
    try:
        import ytmusicapi
    except ImportError:
        print("=" * 60)
        print("Error: 'ytmusicapi' library is not installed.")
        print("Please install it by running:")
        print("  pip install ytmusicapi")
        print("=" * 60)
        sys.exit(1)

check_dependencies()
from ytmusicapi import YTMusic

def load_auth():
    """Finds and returns the path to browser.json or exits with instructions."""
    paths_to_check = [
        'browser.json',
        'scratch/browser.json',
        '../browser.json',
        'd:/01. T3Lab/02 Revit Tools/t3lab-revit-api/browser.json',
        'd:/01. T3Lab/02 Revit Tools/t3lab-revit-api/scratch/browser.json'
    ]
    for p in paths_to_check:
        if os.path.exists(p):
            return p
            
    print("=" * 70)
    print("AUTHENTICATION REQUIRED: 'browser.json' not found!")
    print("=" * 70)
    print("Please follow these steps to connect your YouTube Music account:")
    print("1. Open Chrome/Firefox/Edge and go to https://music.youtube.com")
    print("2. Ensure you are logged in to your account.")
    print("3. Press F12 (Developer Tools) and switch to the 'Network' tab.")
    print("4. In the filter box, type: browse")
    print("5. Refresh the page (F5).")
    print("6. Click on one of the 'browse' requests (type: Fetch/XHR).")
    print("7. Scroll to 'Request Headers', select and copy the ENTIRE block of headers.")
    print("8. Run this command in your terminal to create 'browser.json':")
    print("   ytmusicapi browser")
    print("   (When prompted, paste the headers you copied and press Enter twice).")
    print("=" * 70)
    sys.exit(1)

def tokenize(text):
    """Tokenizes and normalizes text into a set of lowercased alphanumeric words."""
    text = text.lower()
    # Remove common accents/signs or special characters
    text = re.sub(r'[^\w\s]', ' ', text)
    words = text.split()
    
    # Common stopwords to ignore in matching
    stopwords = {
        'music', 'playlist', 'songs', 'best', 'the', 'a', 'an', 'and', 'or', 
        'of', 'to', 'in', 'for', 'with', 'by', 'on', 'at', 'mix', 'album',
        'bài', 'hát', 'nhạc', 'tuyển', 'chọn', 'hay', 'nhất', 'danh', 'sách'
    }
    return [w for w in words if w not in stopwords and len(w) > 1]

def calculate_match_score(song_title, artist_name, playlist_tokens):
    """Calculates how well a song matches the playlist name tokens."""
    if not playlist_tokens:
        return 0.0
        
    song_text = f"{song_title} {artist_name}".lower()
    song_tokens = tokenize(song_text)
    
    # Check overlapping words
    overlap = set(playlist_tokens).intersection(set(song_tokens))
    if not overlap:
        return 0.0
        
    # Score is ratio of matching words to total playlist words
    score = len(overlap) / len(playlist_tokens)
    return score

def sync_playlists(limit_search=100, max_add_per_playlist=10):
    auth_file = load_auth()
    print(f"Initializing YTMusic client with auth file: {auth_file}")
    yt = YTMusic(auth_file)
    
    print("Retrieving library playlists...")
    playlists = yt.get_library_playlists()
    print(f"Found {len(playlists)} playlists in your library.")
    
    for pl in playlists:
        playlist_name = pl['title']
        playlist_id = pl['playlistId']
        
        # Skip special default lists if they don't support adding songs directly
        if playlist_name in ['Liked Music', 'Episodes', 'Your Episodes']:
            continue
            
        print(f"\nProcessing playlist: '{playlist_name}' (ID: {playlist_id})")
        
        # Get existing tracks to avoid duplicates
        try:
            full_playlist = yt.get_playlist(playlist_id, limit=None)
            existing_tracks = set()
            for track in full_playlist.get('tracks', []):
                # Save both videoId and title/artist to check for duplicates
                if track.get('videoId'):
                    existing_tracks.add(track['videoId'])
                # Also index lowercase title to avoid duplicate songs under different videoIds
                existing_tracks.add(track['title'].lower().strip())
        except Exception as e:
            print(f"  Error fetching existing tracks: {e}. Skipping playlist.")
            continue
            
        print(f"  Current tracks in playlist: {len(full_playlist.get('tracks', []))}")
        
        # Tokenize playlist name for keyword matching
        playlist_tokens = tokenize(playlist_name)
        if not playlist_tokens:
            print(f"  Playlist name '{playlist_name}' has no search-eligible keywords. Skipping matching.")
            continue
            
        print(f"  Keywords for matching: {playlist_tokens}")
        print(f"  Searching for up to {limit_search} songs matching '{playlist_name}'...")
        
        try:
            # Search YouTube Music
            search_results = yt.search(query=playlist_name, filter="songs", limit=limit_search)
        except Exception as e:
            print(f"  Error searching songs: {e}")
            continue
            
        candidate_songs = []
        for song in search_results:
            video_id = song.get('videoId')
            if not video_id:
                continue
                
            title = song.get('title', '')
            artists = ", ".join([a['name'] for a in song.get('artists', [])])
            
            # Skip if already exists in playlist (by ID or exact Title)
            if video_id in existing_tracks or title.lower().strip() in existing_tracks:
                continue
                
            score = calculate_match_score(title, artists, playlist_tokens)
            if score > 0:
                candidate_songs.append({
                    'id': video_id,
                    'title': title,
                    'artists': artists,
                    'score': score
                })
                
        # Sort candidates by match score (highest first)
        candidate_songs.sort(key=lambda x: x['score'], reverse=True)
        
        songs_to_add = candidate_songs[:max_add_per_playlist]
        
        if not songs_to_add:
            print("  No new matching songs found.")
            continue
            
        print(f"  Found {len(candidate_songs)} matching candidate songs. Adding top {len(songs_to_add)}:")
        video_ids = []
        for idx, song in enumerate(songs_to_add, 1):
            print(f"    {idx}. [{song['score']:.2f}] {song['title']} - {song['artists']} (ID: {song['id']})")
            video_ids.append(song['id'])
            
        try:
            yt.add_playlist_items(playlist_id, video_ids)
            print(f"  Successfully added {len(video_ids)} songs to playlist '{playlist_name}'.")
        except Exception as e:
            print(f"  Error adding songs to playlist: {e}")
            
        # Polite sleep to respect API limits
        time.sleep(2)

if __name__ == "__main__":
    # Default parameters: search 100 songs, add up to 10 matching ones per playlist
    sync_playlists(limit_search=100, max_add_per_playlist=10)
