# =====================================================
# DAILY MOTIVATION HUB — RICH UI EDITION
# =====================================================

import os
import sys
import time
import random
import datetime
import requests
from pathlib import Path

from rich.console import Console
from rich.panel import Panel
from rich.progress import track
from rich.text import Text

from colorama import Fore, init

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage

# =====================================================
# INIT
# =====================================================

init(autoreset=True)
console = Console()

# =====================================================
# CONFIG
# =====================================================

APIS = {
    "advice": "https://api.adviceslip.com/advice",
    "books": "https://openlibrary.org/search.json?q=motivation",
    "quotes": "https://type.fit/api/quotes",
    "dogs": "https://dog.ceo/api/breeds/image/random"
}


DATA_FILES = {
    "advice": "data/advice.xlsx",
    "books": "data/books.xlsx",
    "quotes": "data/quotes.xlsx",
    "dogs": "data/dogs.xlsx"
}

# =====================================================
# UI HELPERS
# =====================================================

def loading(text="Processing data"):
    for _ in track(range(100), description=f"[green]{text}"):
        time.sleep(0.015)


def show_card(title, message, subtitle="", color="green"):
    panel = Panel(
        message,
        title=f"[bold]{title}[/bold]",
        subtitle=subtitle,
        border_style=color
    )
    console.print(panel)


def now():
    return datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S")


# =====================================================
# SETUP
# =====================================================

def setup():
    Path("data").mkdir(exist_ok=True)
    Path("images").mkdir(exist_ok=True)
    Path("logs").mkdir(exist_ok=True)

    schemas = {
        "advice": ["ID", "Advice", "Time"],
        "books": ["Title", "Author", "Description", "Release", "Pages", "Time"],
        "quotes": ["Quote", "Author", "Category", "Time"],
        "dogs": ["Image URL", "Breed", "Time", "Image"]
    }

    for k, headers in schemas.items():
        if not os.path.exists(DATA_FILES[k]):
            wb = Workbook()
            ws = wb.active
            ws.append(headers)
            wb.save(DATA_FILES[k])


def api(url):
    try:
        loading()
        return requests.get(url, timeout=10).json()
    except Exception as e:
        with open("logs/error.log", "a") as f:
            f.write(f"{now()} | {e}\n")
        return None


# =====================================================
# FEATURES
# =====================================================

def get_advice():
    data = api(APIS["advice"])
    slip = data["slip"]

    show_card(
        title="Advice For Today",
        message=slip["advice"],
        subtitle=f"Advice ID: {slip['id']}",
        color="bright_green"
    )

    wb = load_workbook(DATA_FILES["advice"])
    ws = wb.active
    ws.append([slip["id"], slip["advice"], now()])
    wb.save(DATA_FILES["advice"])

    console.print("[green]✓ Advice saved to Excel[/green]")


def get_quote():
    data = api(APIS["quotes"])

    quote = data.get("quote")
    author = data.get("author", "Unknown")

    show_card(
        title="Inspirational Quote",
        message=f"“{quote}”\n\n— {author}",
        color="yellow"
    )

    wb = load_workbook(DATA_FILES["quotes"])
    ws = wb.active
    ws.append([quote, author, data.get("category", "General"), now()])
    wb.save(DATA_FILES["quotes"])

    console.print("[green]✓ Quote saved to Excel[/green]")


def get_book():
    books = api(APIS["books"])
    book = random.choice(books)

    info = f"""
📘 Title   : {book.get("title")}
✍ Author  : {book.get("author")}
📄 Pages  : {book.get("pages")}
📅 Release: {book.get("releaseDate")}
"""

    show_card(
        title="Harry Potter Recommendation",
        message=info,
        color="cyan"
    )

    wb = load_workbook(DATA_FILES["books"])
    ws = wb.active
    ws.append([
        book.get("title"),
        book.get("author"),
        book.get("description"),
        book.get("releaseDate"),
        book.get("pages"),
        now()
    ])
    wb.save(DATA_FILES["books"])

    console.print("[green]✓ Book saved to Excel[/green]")


def get_dog():
    data = api(APIS["dogs"])
    url = data["message"]

    img_name = f"dog_{int(time.time())}.jpg"
    img_path = f"images/{img_name}"

    with open(img_path, "wb") as f:
        f.write(requests.get(url).content)

    show_card(
        title="Cute Dog Image",
        message=f"🐶 Image downloaded successfully\n\n{img_path}",
        color="magenta"
    )

    wb = load_workbook(DATA_FILES["dogs"])
    ws = wb.active
    row = ws.max_row + 1

    ws.append([url, "Random Dog", now()])
    img = XLImage(img_path)
    img.width = 130
    img.height = 130
    ws.add_image(img, f"D{row}")

    wb.save(DATA_FILES["dogs"])

    console.print("[green]✓ Image saved into Excel[/green]")


def stats():
    text = ""
    total = 0

    for k, path in DATA_FILES.items():
        wb = load_workbook(path)
        ws = wb.active
        count = ws.max_row - 1
        total += count
        text += f"{k.capitalize():8}: {count}\n"

    text += f"\nTOTAL ENTRIES: {total}"

    show_card("Application Statistics", text, color="blue")


def all_data():
    get_advice()
    get_quote()
    get_book()
    get_dog()


# =====================================================
# MENU UI
# =====================================================

def welcome():
    console.print("""
[bold green]
╔══════════════════════════════════════════════════════════════╗
║                     DAILY MOTIVATION HUB                     ║
║         Your one-stop destination for daily inspiration!     ║
╚══════════════════════════════════════════════════════════════╝
[/bold green]
""")

def menu():
    var = """┌─xxxxxxxxxxxx─────────xxxxxxxxxxxxxx────────xxxxxxxx┐
│x┌──────────xxxxxxxxxxx────────────xxxxxxxxxx────┐  │
│ │                                               │  │
│ │      1. Get Advice                            │  │
│ │                                               │  │
│ │      2. Get Quote                             │  │
│ │                                               │  │
│ │      3. Get Harry Potter Recomm.              │  │
│ │                                               │  │
│ │      4. Get Stats                             │  │
│ │                                               │  │
│ │      5. Get Cute Dog Image(mayBe)             │  │
│ │                                               │  │
│ │      6. get all Data                          │  │
│ │                                               │  │
│ │      7. Exit                                  │  │
│ │                                               │  │
│ └───────────────────────────────────────────────┘  │
└────────────────────────────────────────────────────┘"""
    print(Fore.BLUE + var)

# =====================================================
# MAIN
# =====================================================

def main():
    if sys.version_info < (3, 12):
        print("Python 3.12+ required")
        return

    setup()
    welcome()

    while True:
        menu()
        choice = input("Hello, Enter your choice (1-7): 🎉 ")

        if choice == "1":
            get_advice()
        elif choice == "2":
            get_quote()
        elif choice == "3":
            get_book()
        elif choice == "4":
            stats()
        elif choice == "5":
            get_dog()
        elif choice == "6":
            all_data()
        elif choice == "7":
            console.print("\n[bold green]Thank you for using Daily Motivation Hub 🚀[/bold green]")
            break
        else:
            console.print("[bold red]Invalid option![/bold red]")

        input("\nPress Enter to continue...")


if __name__ == "__main__":
    main()
