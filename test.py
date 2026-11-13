# =====================================================
# DAILY MOTIVATION HUB — ULTRA ANIMATED EDITION 🚀
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
from rich.progress import (
    Progress, SpinnerColumn, BarColumn,
    TextColumn, TimeElapsedColumn
)
from rich.align import Align
from rich.live import Live
from rich.table import Table
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
# APIs — STABLE (NO LIMIT)
# =====================================================

APIS = {
    "advice": "https://zenquotes.io/api/random",
    "quotes": "https://type.fit/api/quotes",
    "books": "https://openlibrary.org/search.json?q=harry+potter",
    "dogs": "https://dog.ceo/api/breeds/image/random"
}

DATA_FILES = {
    "advice": "data/advice.xlsx",
    "quotes": "data/quotes.xlsx",
    "books": "data/books.xlsx",
    "dogs": "data/dogs.xlsx"
}

# =====================================================
# UTILITIES
# =====================================================

def now():
    return datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S")


def typewriter(text, speed=0.03):
    for char in text:
        print(char, end="", flush=True)
        time.sleep(speed)
    print()


def startup_animation():
    console.clear()

    with Progress(
        SpinnerColumn(style="cyan"),
        TextColumn("[bold green]{task.description}"),
        BarColumn(),
        TimeElapsedColumn(),
        transient=True
    ) as progress:

        task = progress.add_task("Initializing Daily Motivation Hub", total=100)

        while not progress.finished:
            time.sleep(0.02)
            progress.update(task, advance=2)

    typewriter("✔ Loading UI engine...")
    typewriter("✔ Connecting public APIs...")
    typewriter("✔ Preparing local storage...")
    time.sleep(0.4)


def animated_loading(text="Fetching data"):
    with Progress(
        SpinnerColumn(style="green"),
        TextColumn("[bold cyan]{task.description}"),
        BarColumn(),
        transient=True
    ) as progress:
        t = progress.add_task(text, total=100)
        while not progress.finished:
            time.sleep(0.015)
            progress.update(t, advance=3)


def show(title, message, color="green"):
    console.print(
        Panel(
            Align.center(message),
            title=f"[bold]{title}[/bold]",
            border_style=color,
            padding=(1, 2)
        )
    )

# =====================================================
# SETUP
# =====================================================

def setup():
    Path("data").mkdir(exist_ok=True)
    Path("images").mkdir(exist_ok=True)

    schemas = {
        "advice": ["ID", "Advice", "Time"],
        "quotes": ["Quote", "Author", "Time"],
        "books": ["Title", "Author", "Year", "Pages", "Time"],
        "dogs": ["Image URL", "Breed", "Time", "Image"]
    }

    for key, headers in schemas.items():
        if not os.path.exists(DATA_FILES[key]):
            wb = Workbook()
            ws = wb.active
            ws.append(headers)
            wb.save(DATA_FILES[key])


def api(url):
    try:
        animated_loading("Contacting API")
        r = requests.get(url, timeout=10)
        if r.status_code != 200:
            return None
        return r.json()
    except:
        return None

# =====================================================
# FEATURES
# =====================================================

def get_advice():
    data = api(APIS["advice"])
    if not data:
        show("Advice", "API unavailable", "red")
        return

    q = data[0]
    advice = q.get("q", "Stay positive.")
    author = q.get("a", "Unknown")

    show("💡 Daily Advice", f"{advice}\n\n— {author}", "bright_green")

    wb = load_workbook(DATA_FILES["advice"])
    ws = wb.active
    ws.append(["-", advice, now()])
    wb.save(DATA_FILES["advice"])


def get_quote():
    data = api(APIS["quotes"])
    if not data:
        show("Quote", "API unavailable", "red")
        return

    q = random.choice(data)
    text = q.get("text", "")
    author = q.get("author") or "Unknown"

    show("✨ Inspirational Quote", f"{text}\n\n— {author}", "yellow")

    wb = load_workbook(DATA_FILES["quotes"])
    ws = wb.active
    ws.append([text, author, now()])
    wb.save(DATA_FILES["quotes"])


def get_book():
    data = api(APIS["books"])
    if not data or "docs" not in data:
        show("Books", "No books found", "red")
        return

    b = random.choice(data["docs"])

    title = b.get("title", "Unknown")
    author = ", ".join(b.get("author_name", ["Unknown"]))
    year = b.get("first_publish_year", "N/A")
    pages = b.get("number_of_pages_median", "N/A")

    msg = (
        f"📘 Title : {title}\n"
        f"✍ Author: {author}\n"
        f"📅 Year  : {year}\n"
        f"📄 Pages : {pages}"
    )

    show("📚 Harry Potter Recommendation", msg, "cyan")

    wb = load_workbook(DATA_FILES["books"])
    ws = wb.active
    ws.append([title, author, year, pages, now()])
    wb.save(DATA_FILES["books"])


def get_dog():
    data = api(APIS["dogs"])
    if not data:
        show("Dog", "API unavailable", "red")
        return

    url = data["message"]
    img_name = f"dog_{int(time.time())}.jpg"
    path = f"images/{img_name}"

    img = requests.get(url).content
    with open(path, "wb") as f:
        f.write(img)

    show("🐶 Cute Dog Image", f"Saved to:\n{path}", "magenta")

    wb = load_workbook(DATA_FILES["dogs"])
    ws = wb.active
    row = ws.max_row + 1
    ws.append([url, "Random Dog", now()])

    pic = XLImage(path)
    pic.width = 120
    pic.height = 120
    ws.add_image(pic, f"D{row}")
    wb.save(DATA_FILES["dogs"])


def stats():
    table = Table(title="📊 Live Statistics", style="cyan")
    table.add_column("Category")
    table.add_column("Entries")

    total = 0

    for k, file in DATA_FILES.items():
        wb = load_workbook(file)
        ws = wb.active
        count = ws.max_row - 1
        total += count
        table.add_row(k.capitalize(), str(count))

    table.add_row("TOTAL", str(total))
    console.print(table)


def all_data():
    get_advice()
    get_quote()
    get_book()
    get_dog()

# =====================================================
# UI
# =====================================================

def banner():
    console.print("""
[bold green]
╔══════════════════════════════════════════════════════════════╗
║                     DAILY MOTIVATION HUB                     ║
║         Your one-stop destination for daily inspiration!     ║
╚══════════════════════════════════════════════════════════════╝
[/bold green]
""")


def menu():
    print(Fore.CYAN + """
┌──────────────────────────────────────────────────────────────┐
│  1. Get Advice                                                │
│  2. Get Quote                                                 │
│  3. Get Harry Potter Recommendation                           │
│  4. View Stats                                                │
│  5. Get Cute Dog Image                                        │
│  6. Get All Data                                              │
│  7. Exit                                                      │
└──────────────────────────────────────────────────────────────┘
""")


# =====================================================
# MAIN
# =====================================================

def main():
    if sys.version_info < (3, 10):
        print("Python 3.10+ required")
        return

    setup()
    startup_animation()
    banner()

    while True:
        menu()
        ch = input("🎯 Enter choice (1–7): ")

        if ch == "1":
            get_advice()
        elif ch == "2":
            get_quote()
        elif ch == "3":
            get_book()
        elif ch == "4":
            stats()
        elif ch == "5":
            get_dog()
        elif ch == "6":
            all_data()
        elif ch == "7":
            console.print("\n[bold green]Thank you for using Daily Motivation Hub 🚀[/bold green]")
            break
        else:
            console.print("[red]Invalid choice[/red]")

        input("\nPress Enter to continue...")


if __name__ == "__main__":
    main()
