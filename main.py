# This cmd program scraps a Youtube playlist
# and writes the name, artist/uploder and link
# into an new or existing Excel file

# for song playlists replace the "www"
# in the playlists link with "music"
# to get better results

# necessary libraries: bs4, openpyxl, requests
from bs4 import BeautifulSoup
import openpyxl
import requests


def check_input(url):
    match url:
        case s if s.startswith("https://www.youtube.com/playlist?list="):
            print("Valid Youtube link, continuing...")
            return 0
        case s if s.startswith("https://music.youtube.com/playlist?list="):
            print("Valid Youtube-Music link, continuing...")
            return 1
        case "-h":
            print("""
            This cmd program scraps a Youtube playlist
            and writes the name, artist/uploder and link
            into an new or existing Excel file.
    
            Only works with links like this: "https://www.youtube.com/playlist?list=".
            For better results with song playlists use a YT-music
            playlist like this: "https://music.youtube.com/playlist?list=". \n
                    """)
            return 8
        case _:
            print("\nInvalid input. Please try again.\n")
            return 9

def get_normal_info():
    pass

def get_music_info():
    pass

def main():
    print("Welcome to YT-Playlist2Excel")
    while True:
        #url = input("Please input your playlist link or -h for help: ")
        # Test URL
        url = "https://music.youtube.com/playlist?list=PLdxfyzVbTnHn_7pgh4twqXGQva2EgfnF7"

        mode = check_input(url)
        if mode == 0 or mode == 1:
            break

    print("Finished")
    input("Press any key to close...")



if __name__ == "__main__":
    main()
