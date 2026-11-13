import os
from rich.text import Text
import time
import json
import datetime
import requests
import sys
import random
from pathlib import Path
from  openpyxl import load_workbook, Workbook
from rich.progress import track
from time import sleep
from colorama import Fore, init

from rich.progress import track
from time import sleep
from colorama import Fore, init
from rich.console import Console
from rich.panel import Panel
advice = {"id":0, "advice":""}
task = {"task":""}
init(autoreset=True)
console = Console()
DATA_FILES = {
    'advice': 'data/advice_data.xlsx',
    'books': 'data/books_data.xlsx',
    'quotes': 'data/quotes_data.xlsx',
    'dogs': 'data/dogs_data.xlsx'
}
# Global configuration
APIS = {
    'advice': 'https://api.adviceslip.com/advice',
    'books': 'https://potterapi-fedeperin.vercel.app/en/books',
    'quotes': 'https://hindi-quotes.vercel.app/random',
    'dogs': 'https://dog.ceo/api/breeds/image/random'
}


def welcome():   
    var = """╔══════════════════════════════════════════════════════════════╗
║                     DAILY MOTIVATION HUB                     ║
║                                                              ║
║         Your one-stop destination for daily inspiration!     ║
║                                                              ║
║   Books  |   Advice  |   Quotes  |   Cute Dogs               ║
╚══════════════════════════════════════════════════════════════╝"""
    print(Fore.GREEN + var)

def menu():
    var = """┌─xxxxxxxxxxxx─────────xxxxxxxxxxxxxx────────xxxxxxxx┐
│x┌──────────xxxxxxxxxxx────────────xxxxxxxxxx────┐  │
│ │                                               │  │
│ │     💡 1. Get Advice                         │  │
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

def setup_directories():
    """Create necessary directories if they don't exist"""
    if 1== 2:
        if not os.path.exists("data"):
            os.mkdir("data")
    Path('data').mkdir(exist_ok=True)
    Path('logs').mkdir(exist_ok=True)

def init_excel_files():
    """Initialize Excel files with headers if they don't exist"""
    file_schemas = {
        'advice': ['ID', 'Advice', 'Timestamp'],
        'books': ['Title', 'Author', 'Description', 'Release Date', 'Pages', 'Timestamp'],
        'quotes': ['Quote', 'Author', 'Category', 'Timestamp'],
        'dogs': ['Image URL', 'Breed Info', 'Timestamp']
    }
    
    for file_type, headers in file_schemas.items():
        file_path = DATA_FILES[file_type]
        if not os.path.exists(file_path):
            wb = Workbook()
            ws = wb.active
            ws.title = f"{file_type.capitalize()} Data"
            
            # # Add headers with styling
            # for col, header in enumerate(headers, 1):
            #     cell = ws.cell(row=1, column=col, value=header)
            #     cell.font = Font(bold=True, color="FFFFFF")
            #     cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            #     cell.alignment = Alignment(horizontal="center")
            
            wb.save(file_path)
def process_data():
    sleep(0.02)

def show_loading():
    for _ in track(range(100), description='[green]Processing data'):
        process_data()
def time_now():
    x = datetime.datetime.now()
    return x.strftime("%c")

def log_error(error_message):
    """Log errors to file"""
    with open('logs/app.log', 'a') as log_file:
        log_file.write(f"[{time_now()}] ERROR: {error_message}\n")


def make_api_request(api_name, timeout=10):
    """Make API request with error handling"""
    try:
        show_loading()
        response = requests.get(APIS[api_name], timeout=timeout)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        log_error(f"API Error for {api_name}: {str(e)}")
        return None
    except json.JSONDecodeError as e:
        log_error(f"JSON Decode Error for {api_name}: {str(e)}")
        return None
def save_to_excel(data_type, data):
    """Save data to Excel file"""
    try:
        file_path = DATA_FILES[data_type]
        wb = load_workbook(file_path)
        ws = wb.active
        # Get next row
        next_row = ws.max_row + 1
        # Save based on data type
        if data_type == 'advice':
            ws.append([advice["id"],advice["advice"],time_now()])
        elif data_type == 'books':
            # Save a random book from the list
            book = random.choice(data)
            #HOMEWORK Change below to APPEND 
            # ws.cell(row=next_row, column=1, value=book.get('title', 'N/A'))
            # ws.cell(row=next_row, column=2, value=book.get('author', 'N/A'))
            # ws.cell(row=next_row, column=3, value=book.get('description', 'N/A'))
            # ws.cell(row=next_row, column=4, value=book.get('releaseDate', 'N/A'))
            # ws.cell(row=next_row, column=5, value=book.get('pages', 'N/A'))
            # ws.cell(row=next_row, column=6, value=time_now())
        
        elif data_type == 'quotes':
            pass
            #HOMEWORK Change below to APPEND 
            # ws.cell(row=next_row, column=1, value=data.get('quote', 'N/A'))
            # ws.cell(row=next_row, column=2, value=data.get('author', 'Anonymous'))
            # ws.cell(row=next_row, column=3, value=data.get('category', 'General'))
            # ws.cell(row=next_row, column=4, value=time_now())
        
        elif data_type == 'dogs':
            #HOMEWORK Change below to APPEND 
            pass
            # ws.cell(row=next_row, column=1, value=data.get('message', 'N/A'))
            # ws.cell(row=next_row, column=2, value='Random Dog Image')
            # ws.cell(row=next_row, column=3, value=time_now())
        
        wb.save(file_path)
        return True
    except Exception as e:
        log_error(f"Excel Save Error for {data_type}: {str(e)}")
        return False

def get_advice():
    """Fetch and display random advice"""
    data = make_api_request('advice')
    var = data["slip"]
    id_val = "Advice Id:" + str(var["id"])
    advice_val = var["advice"]
    card(title="Advice For Today",text=advice_val,color="GREEN",subtitle=id_val)
    advice["id"] = id_val
    advice["advice"] = advice_val    
    if save_to_excel('advice', data):
        print(f"{Fore.GREEN}✓ Advice saved to Excel!")
    else:
        print(f"{Fore.RED}✗ Failed to save advice")

def card(text,title,subtitle, color):
    panel = Panel(f"{text}",style=color, title=title, subtitle=subtitle)
    console.print(panel)
def get_book():
    """Fetch and display random book recommendation"""
    data = make_api_request('books')
    book = random.choice(data)
        
    book_info = {
        "Title": book.get('title', 'Unknown'),
        "Author": book.get('author', 'Unknown'),
        "Pages": book.get('pages', 'Unknown'),
        "Release Date": book.get('releaseDate', 'Unknown')
    }
        
    #card(type="book",title="📚 BOOK RECOMMENDATION",text=book_info,color="CYAN",subtitle="Enjoy Reading..")    
    card(title="📚 BOOK RECOMMENDATION",text=book_info,color="CYAN",subtitle="Enjoy Reading..")    
    # Show description separately if available
    if book.get('description'):
        #card(type="book",title="📖 DESCRIPTION",text=book['description'],color="GREEN",subtitle="Desc")
        card(title="📖 DESCRIPTION",text=book['description'],color="GREEN",subtitle="Desc")
        
        if save_to_excel('books', data):
            print(f"{Fore.GREEN}✓ Book recommendation saved to Excel!")
        else:
            print(f"{Fore.RED}✗ Failed to save book recommendation")

def get_quote():
    """Fetch and display inspirational quote"""
    data = make_api_request('quotes')
    if data:
        quote_text = data.get('quote', 'No quote available')
        author = data.get('author', 'Anonymous')
        #HOMEWORK DO A CARD LAYOUT BELOW 
        #card("🎯 INSPIRATIONAL QUOTE", f'"{quote_text}"\n\n- {author}', Fore.YELLOW)
        if save_to_excel('quotes', data):
            print(f"{Fore.GREEN}✓ Quote saved to Excel!")
        else:
            print(f"{Fore.RED}✗ Failed to save quote")
    else:
        print(f"{Fore.RED}❌ Failed to fetch quotes. Please try again later.")

def get_dog_image():
    """Fetch and display dog image URL"""
    data = make_api_request('dogs')
    if data and data.get('status') == 'success':
        image_url = data['message']
        #HOMEWORK ALSO DOWNLOAD DOG IMAGE IN A FOLDER CALLED IMAGES 
        
        # response = requests.get(image_url)
        # with open("abc.jpeg", "wb") as f:
        #     f.write(response.content)
        # Extract breed info from URL if possible
        
        breed_info = "Random Dog"
        # if '/breeds/' in image_url:
        #     try:
        #         breed_part = image_url.split('/breeds/')[1].split('/')[0]
        #         if '-' in breed_part:
        #             sub_breed, main_breed = breed_part.split('-')
        #             breed_info = f"{sub_breed.title()} {main_breed.title()}"
        #         else:
        #             breed_info = breed_part.title()
        #     except:
        #         pass
        
        dog_info = {
            "Breed": breed_info,
            "Image URL": image_url,
            "Note": "Copy the URL to your browser to see the cute dog!"
        }
        #HOMEWORK
        #card("🐕 CUTE DOG IMAGE", dog_info, "GREEN)
        
        if save_to_excel('dogs', data):
            print(f"{Fore.GREEN}✓ Dog image info saved to Excel!")
        else:
            print(f"{Fore.RED}✗ Failed to save dog image info")
    else:
        print(f"{Fore.RED}❌ Failed to fetch dog images. Please try again later.")

def get_everything():
    """Fetch all content types - daily dose of motivation!"""
    print(f"\n{Fore.CYAN}🌟 Preparing your daily dose of motivation...\n")
    
    print(f"{Fore.YELLOW}{'='*60}")
    get_advice()
    
    print(f"\n{Fore.YELLOW}{'='*60}")
    get_quote()
    
    print(f"\n{Fore.YELLOW}{'='*60}")
    get_book()
    
    print(f"\n{Fore.YELLOW}{'='*60}")
    get_dog_image()
    
    print(f"\n{Fore.GREEN}🎉 Your daily motivation dose is complete! Have a great day! 🎉")

def show_statistics():
    """Display statistics about saved data"""
    stats = {}
    total_entries = 0
    
    for data_type, file_path in DATA_FILES.items():
        try:
            if os.path.exists(file_path):
                wb = load_workbook(file_path)
                ws = wb.active
                count = ws.max_row - 1  # Subtract header row
                stats[data_type] = max(0, count)
                total_entries += stats[data_type]
            else:
                stats[data_type] = 0
        except:
            stats[data_type] = 0
    
    stats_info = {
        "📝 Total Entries": total_entries,
        "💡 Advice Count": stats['advice'],
        "📚 Books Count": stats['books'],
        "🎯 Quotes Count": stats['quotes'],
        "🐕 Dogs Count": stats['dogs']
    }
    
    #card("📊 DATA STATISTICS", stats_info, Fore.CYAN)

def main():
    """Main application function"""
    # Check Python version
    if sys.version_info < (3, 12):
        print("Error: This application requires Python 3.12 or higher")
        sys.exit(1)
    
    # Setup application
    setup_directories()
    init_excel_files()
    
    # Display welcome
    welcome()
    print("\n\n")
    
    
    # Main application loop
    while True:
        try:
            menu()
            console.print("Hello, [bold magenta]Enter your choice (1-7): [/bold magenta]! :tada:")
            choice = input().strip()
            
            if choice == '1':
                get_advice()
            elif choice == '2':
                get_book()
            elif choice == '3':
                get_quote()
            elif choice == '4':
                get_dog_image()
            elif choice == '5':
                get_everything()
            elif choice == '6':
                show_statistics()
            elif choice == '7':
                print(f"\n{Fore.GREEN}Thank you for using Daily Motivation Hub! 🌟")
                print(f"Stay motivated and have a great day! 💪")
                break
            else:
                t = Text("Warning: Invalid Choice")
                t.stylize("bold red")
                t.append(" — Only 1-7 Allowed!", style="italic yellow")
                console.print(t)
                #print(f"{Fore.RED}❌ Invalid choice. Please enter a number between 1-7.")
            
            # Wait for user to continue
            if choice in ['1', '2', '3', '4', '5', '6']:
                input(f"\n{Fore.LIGHTBLACK_EX}Press Enter to continue...")
                print("\n" * 2)  # Clear some space
            
        except KeyboardInterrupt:
            print(f"\n\n{Fore.YELLOW}👋 Goodbye! Thanks for using Daily Motivation Hub!")
            break
        except Exception as e:
            print(f"{Fore.RED}❌ An unexpected error occurred: {str(e)}")
            log_error(f"Unexpected error in main loop: {str(e)}")

if __name__ == "__main__":
    main()