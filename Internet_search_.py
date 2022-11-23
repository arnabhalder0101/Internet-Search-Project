
import wikipedia
import wikipedia as wk
from win32com.client import Dispatch
# from spellchecker import SpellChecker   # spellchecker currently not in use
from datetime import datetime
from googlesearch import search
from bing_image_urls import bing_image_urls

# spell = SpellChecker()


def speak(str2):
    speak1 = Dispatch("SAPI.Spvoice")
    speak1.speak(str2)


def image_url_search(query, limit):
    print("Few image urls that matches with your search :")
    for urls in bing_image_urls(query, limit=limit):
        print(urls)
    return 0


def search_with_google(query, number):
    print("Those might be helpful --")
    for google_result in search(query, lang="en", num=number, stop=number, pause=3):
        print(google_result)
    return 0


def search_me(string1, integer1):
    print(f"\tsearching for :\t{string1}\n")
    print(f"You can also search for :\t{wikipedia.search(string1)} \n")
    image_url_search(string1, integer1)
    # print(f"You can also search for {wikipedia.summary(wikipedia.suggest(string1))} ")
    try:
        result = wk.summary(string1, sentences=integer1)
        print(f"Your searched results : \t{result}\n")
        return 0

    except:
        search_with_google(string1, integer1)

    finally:
        with open("Search_history.txt", "a") as f:
            f.write(f"Searched for -'{string1}' at {datetime.now()}\n")
        user_inp1 = input("\tPress 'h' to hear the results with Arkish :\n\tPress 'n' to ignore this feature :\n")
        if user_inp1 == "h":
            print("My bot 'Arkish' is Reading for you...")
            try:
                speak(result)
            except:
                speak_this = "These are Links, it's worthless to read. Click it, if you want results."
                speak(speak_this)
        elif user_inp1 == "n":
            print("Your order is my command. Ignoring...")


if __name__ == '__main__':
    while True:

        str_ = input("Enter what you're looking for :")
        num_ = input("Enter number of sentences/link of results you want :")
        result1 = search_me(str_.capitalize(), int(num_))

        user_inp2 = input("\tType 'History' to read Search results; Type 'i' to Ignore\n")
        if user_inp2 == "History":
            with open("Search_history.txt", "r") as f1:
                print(f1.read())
        elif user_inp2 == "i":
            print("Thanks for confirming, ignoring...")

        user_inp3 = input("\tPress 'c' to continue & 'q' to quit\n ")

        if user_inp3 == "c":
            continue
        else:
            break
