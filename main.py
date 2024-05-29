# This cmd program scraps a Youtube playlist
# and writes the name, artist/uploder and link
# into an new or existing Excel file

# for song playlists replace the "www"
# in the playlists link with "music"
# to get better results

# necessary libraries: bs4, openpyxl, requests
import openpyxl
from sys import exit
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep


def check_input(url):
    match url:
        case s if s.startswith("https://www.youtube.com/playlist?list="):
            print("Valid Youtube link, continuing...")
            print()
            return 0
        case s if s.startswith("https://music.youtube.com/playlist?list="):
            print("Valid Youtube-Music link, continuing...")
            print()
            return 1
        case "-q":
            print("Quitting the program.\nGoodbye...")
            exit(1)
        case "-h":
            print("""
            This cmd program scraps a Youtube playlist
            and writes the name, artist/uploder and link
            into an new or existing Excel file.
    
            Only works with links like this: "https://www.youtube.com/playlist?list=".
            Or YT-music playlists like this: "https://music.youtube.com/playlist?list=".
             
             My commands like -h for help and -q to quit are case sensitive.
             I don't feel like implementing that rn ngl.
             
             I'm just having fun with this and tr
             Thank you for using my program :3\n
            """)
            return 8
        case _:
            print("\nInvalid input. Please try again or press -h for help.\n")
            return 9


def get_normal_info(driver):
    # Cookie & load complete page
    driver.find_element(By.XPATH, '/html/body/c-wiz/div/div/div/div[2]/div[1]/div[3]/div[1]/form[2]/div').click()
    # driver.find_element(By.TAG_NAME, "body").send_keys(Keys.CONTROL+Keys.END)
    html = driver.find_element(By.TAG_NAME, 'html')
    while True:
        s1 = driver.page_source
        html.send_keys(Keys.END)
        sleep(1)
        html.send_keys(Keys.END)
        s2 = driver.page_source
        if s1 == s2:
            break

    # The thing that actually does stuff:

    content = driver.find_element(By.ID, "contents")

    results = content.find_elements(By.CLASS_NAME, "style-scope.ytd-playlist-video-list-renderer")

    vids = results[4:(len(results) - 1)]

    print("-" * 20)

    entries = []

    for entry in vids:
        try:
            new_entry = []

            entry_text = entry.text.splitlines()
            new_entry.append(entry_text[2:4])

            try:
                link_pos = entry.find_element(By.CLASS_NAME, "yt-simple-endpoint.style-scope.ytd-playlist-video-renderer")
                link = link_pos.get_attribute('href').split('&')
                new_entry.append(link[0])

            except:
                new_entry.append("unavailable")

            print(new_entry)
            entries.append(new_entry)

        except Exception as error:
            new_entry = [error, "unavailable", "unavailable"]

    return entries


def get_music_info(driver):
    # Cookie & load complete page
    driver.find_element(By.XPATH, '/html/body/c-wiz/div/div/div/div[2]/div[1]/div[3]/div[1]/form[2]/div').click()
    #driver.find_element(By.TAG_NAME, "body").send_keys(Keys.CONTROL+Keys.END)
    html = driver.find_element(By.TAG_NAME, 'html')
    while True:
        s1 = driver.page_source
        sleep(1)
        html.send_keys(Keys.PAGE_DOWN)
        html.send_keys(Keys.END)
        sleep(1)
        html.send_keys(Keys.END)
        s2 = driver.page_source
        if s1 == s2:
            break

    # The thing that actually does stuff:

    content = driver.find_element(By.ID, "contents")

    results = content.find_elements(By.CLASS_NAME, "style-scope.ytmusic-playlist-shelf-renderer")

    vids = results[2:(len(results) - 1)]

    print("-" * 20)

    entries = []

    for entry in vids:
        try:
            new_entry = []

            entry_text = entry.text.splitlines()
            new_entry.append(entry_text[0:2])

            try:
                link_pos = entry.find_element(By.CLASS_NAME, "yt-simple-endpoint.style-scope.yt-formatted-string")
                link = link_pos.get_attribute('href').split('&')
                new_entry.append(link[0])

            except:
                new_entry.append("unavailable")

            print(new_entry)
            entries.append(new_entry)

        except Exception as error:
            new_entry = [error, "unavailable", "unavailable"]

    return entries


def search_for_point(list):
    point = "•"

    for i in range(len(list)):
        if list[i] == point:
            return [(i - 2), i]



def get_informations(driver, mode):

    pathToResults = ["style-scope.ytd-playlist-video-list-renderer", "style-scope.ytmusic-playlist-shelf-renderer"]
    startOfActualResults = [4, 2]
    actualLinkPos = ["yt-simple-endpoint.style-scope.ytd-playlist-video-renderer", "yt-simple-endpoint.style-scope.yt-formatted-string"]


    # Cookie & load complete page
    driver.find_element(By.XPATH, '/html/body/c-wiz/div/div/div/div[2]/div[1]/div[3]/div[1]/form[2]/div').click()
    # driver.find_element(By.TAG_NAME, "body").send_keys(Keys.CONTROL+Keys.END)
    html = driver.find_element(By.TAG_NAME, 'html')
    while True:
        s1 = driver.page_source
        sleep(1)
        html.send_keys(Keys.PAGE_DOWN)
        html.send_keys(Keys.END)
        sleep(1)
        html.send_keys(Keys.END)
        s2 = driver.page_source
        if s1 == s2:
            break

    # The thing that actually does stuff:

    content = driver.find_element(By.ID, "contents")

    results = content.find_elements(By.CLASS_NAME, pathToResults[mode])

    vids = results[startOfActualResults[mode]:(len(results) - 1)]

    print("-" * 20)

    entries = []

    for entry in vids:
        try:
            print("-" * 20)
            print(entry.text)
            new_entry = []

            entry_text = entry.text.splitlines()

            match mode:
                case 0:
                    actualEntryText = search_for_point(entry_text)
                    new_entry.append(entry_text[actualEntryText[0]:actualEntryText[1]])
                case 1:
                    new_entry.append(entry_text[0:2])

            try:
                link_pos = entry.find_element(By.CLASS_NAME, actualLinkPos[mode])
                link = link_pos.get_attribute('href').split('&')
                new_entry.append(link[0])

            except:
                new_entry.append("unavailable")

            print(new_entry)
            entries.append(new_entry)

        except Exception as error:
            entries = [error, "unavailable", "unavailable"]

    return entries

def main():
    print("Welcome to YT-Playlist2Excel")
    while True:
        #url = input("Please input your playlist link, press -h for help or -q to quit: ")
        # Test URL
        #url = "https://www.youtube.com/playlist?list=PLdxfyzVbTnHn_7pgh4twqXGQva2EgfnF7"
        url = "https://www.youtube.com/playlist?list=PLdxfyzVbTnHl3sJ5mhYTULvKlj4q0HIDY"

        mode = check_input(url)
        if mode < 5:
            break

    # Init Selenium
    driver = webdriver.Firefox()
    driver.get(url)
    driver.implicitly_wait(10)

    # Getting the information

    entries = get_informations(driver, mode)

    print(entries)
    print(str(len(entries)) + " Videos found.")
    driver.close()

    # Save in Excel

    print("Finished...")


if __name__ == "__main__":
    main()
