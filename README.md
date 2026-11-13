# 🌟 Daily Motivation Hub

### A Python-Powered Application for Daily Inspiration Through Public APIs

![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)
![API Integration](https://img.shields.io/badge/API-Integration-009688?style=for-the-badge)
![Excel](https://img.shields.io/badge/Excel-Automation-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Project](https://img.shields.io/badge/Project-Python%20Development-555555?style=for-the-badge)

**One application. Multiple APIs. Your daily dose of inspiration.**

---

## 📌 About The Project

**Daily Motivation Hub** is a Python-based application that integrates multiple public APIs to provide users with daily advice, inspirational quotes, book recommendations and random dog images.

The application combines API integration, JSON data processing, interactive terminal interfaces and Excel-based data storage into one centralized platform.

Built using Python, the application uses Rich and Colorama to create an interactive terminal experience and OpenPyXL to manage locally stored information.

The main objective is to explore practical Python development, REST API integration, data handling and automation.

---

## ✨ Features

| Feature | Description |
|:---|:---|
| 💡 Daily Advice | Fetch random advice from a public API |
| ✨ Inspirational Quotes | Retrieve inspirational quotes and authors |
| 📚 Book Recommendations | Discover books using Open Library |
| 🐶 Random Dog Images | Fetch and download random dog images |
| 📊 Application Statistics | View stored information and records |
| ⚡ Get Everything | Retrieve all available content categories |
| 📁 Excel Data Storage | Store retrieved information in Excel workbooks |
| 🎨 Interactive Terminal | Colorful terminal interface and progress animations |
| 🖼️ Image Management | Download and organize dog images locally |

---

## 🛠️ Technology Stack

### 💻 Programming Language

![Python](https://skillicons.dev/icons?i=python)

**Python** – Core programming language used to develop the application.

### 📦 Python Libraries

![Requests](https://img.shields.io/badge/Requests-HTTP%20Requests-2CA5E0?style=flat-square)
![Rich](https://img.shields.io/badge/Rich-Terminal%20UI-555555?style=flat-square)
![Colorama](https://img.shields.io/badge/Colorama-Terminal%20Colors-008000?style=flat-square)
![OpenPyXL](https://img.shields.io/badge/OpenPyXL-Excel%20Automation-217346?style=flat-square)

| Library | Purpose |
|:---|:---|
| Requests | Sending HTTP requests to external APIs |
| Rich | Interactive terminal panels, tables and progress displays |
| Colorama | Colored terminal output |
| OpenPyXL | Creating, reading and updating Excel workbooks |
| Pathlib | Managing files and directories |
| Datetime | Working with timestamps |

---

## 🌐 API Integrations

The application integrates four public API services.

| API | Function | Website |
|:---|:---|:---|
| Advice API | Random daily advice | [ZenQuotes](https://zenquotes.io/) |
| Type.fit | Inspirational quotes | [Type.fit](https://type.fit/api/quotes) |
| Open Library | Book recommendations | [Open Library](https://openlibrary.org/) |
| Dog CEO | Random dog images | [Dog CEO](https://dog.ceo/dog-api/) |

### How API Integration Works

1. The application sends an HTTP request to the selected API.
2. The API returns a JSON response.
3. Python processes the returned information.
4. Rich displays the information in the terminal.
5. OpenPyXL stores supported records in Excel workbooks.
6. Downloaded images are saved in the local images directory.

---

## 🖥️ Application Menu

The application provides an interactive menu with the following options:

| Option | Function |
|:---:|:---|
| 1 | Get Daily Advice |
| 2 | Get Inspirational Quote |
| 3 | Get Book Recommendation |
| 4 | View Application Statistics |
| 5 | Get Random Dog Image |
| 6 | Get All Data |
| 7 | Exit Application |

---

## 📂 Project Structure

```text
api-hub/
│
├── data/
│   ├── advice.xlsx
│   ├── books.xlsx
│   ├── quotes.xlsx
│   └── dogs.xlsx
│
├── images/
│   └── dog_*.jpg
│
├── api-hub-final.py
├── finale_main (1).py
├── test.py
├── requirements.txt
├── .gitignore
└── README.md
```

### Folder Description

| File / Folder | Description |
|:---|:---|
| `data/` | Stores Excel workbooks and retrieved information |
| `images/` | Stores downloaded dog images |
| `api-hub-final.py` | Alternative application implementation |
| `finale_main (1).py` | Earlier application implementation |
| `test.py` | Interactive application implementation |
| `requirements.txt` | Python dependency list |
| `.gitignore` | Files and folders excluded from Git |
| `README.md` | Project documentation |

---

## ⚙️ Installation & Setup

Follow these instructions to run the project on your local machine.

### 1. Clone the Repository

```bash
git clone https://github.com/umbaupesrehaan1-art/api-hub.git
```

### 2. Navigate to the Project Directory

```bash
cd api-hub
```

### 3. Check Python Installation

Make sure Python 3.10 or higher is installed.

```bash
python --version
```

### 4. Create a Virtual Environment (Recommended)

```bash
python -m venv venv
```

Activate it on Windows:

```powershell
venv\Scripts\activate
```

### 5. Install Dependencies

```bash
pip install requests rich colorama openpyxl
```

Alternatively, if you have a `requirements.txt` file:

```bash
pip install -r requirements.txt
```

### 6. Run the Application

```bash
python test.py
```

The interactive application menu will appear in your terminal.

---

## 📊 Data Storage

The application uses Excel workbooks for local data storage.

| Workbook | Information |
|:---|:---|
| `advice.xlsx` | Advice text and related records |
| `quotes.xlsx` | Inspirational quotes and authors |
| `books.xlsx` | Book details and related information |
| `dogs.xlsx` | Dog image information and related records |

Downloaded images are maintained in the `images/` directory.

---

## 🔍 Learning Outcomes

Working on this project helped me explore and practise:

- Python application development.
- Integrating multiple external REST APIs.
- HTTP requests and JSON response handling.
- Working with external API data.
- Excel automation using OpenPyXL.
- File and directory management.
- Building interactive command-line applications.
- Terminal styling and progress animations.
- Combining multiple Python libraries in one project.

---

## 🚀 Future Improvements

Some improvements I would like to explore:

- [ ] Improve API error handling and response validation.
- [ ] Introduce a modular application architecture.
- [ ] Add structured logging.
- [ ] Expand automated test coverage.
- [ ] Integrate additional public APIs.
- [ ] Add a graphical user interface.
- [ ] Improve data export and reporting.
- [ ] Add more advanced data management features.

---

## 👨‍💻 Developer

### Rehaan Usama Umba

**Computer Engineering Student | AI & Software Developer**

Interested in Artificial Intelligence, Python development, API integration and building practical software solutions.

📍 Mumbai, Maharashtra, India

📧 **Email:** [umbaupesrehaan1@gmail.com](mailto:umbaupesrehaan1@gmail.com)

🐙 **GitHub:** [umbaupesrehaan1-art](https://github.com/umbaupesrehaan1-art)

---

<p align="center">
  **Learn • Build • Experiment • Improve**
</p>

<p align="center">
  ⭐ Thank you for visiting Daily Motivation Hub!
</p>
