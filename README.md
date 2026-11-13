
# 🌟 Daily Motivation Hub

### A Python-Powered Application for Daily Inspiration Through Public APIs

\<p align="center">
&#x20; \<img src="[https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&amp;logo=python&amp;logoColor=white](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge\&logo=python\&logoColor=white)"/>
&#x20; \<img src="[https://img.shields.io/badge/API-Integration-009688?style=for-the-badge](https://img.shields.io/badge/API-Integration-009688?style=for-the-badge)"/>
&#x20; \<img src="[https://img.shields.io/badge/Excel-Automation-217346?style=for-the-badge&amp;logo=microsoftexcel&amp;logoColor=white](https://img.shields.io/badge/Excel-Automation-217346?style=for-the-badge\&logo=microsoftexcel\&logoColor=white)"/>
&#x20; \<img src="[https://img.shields.io/badge/Project-Student%20Development-555555?style=for-the-badge](<https://img.shields.io/badge/Project-Student%20Development-555555?style=for-the-badge>)"/>
\</p>

\<p align="center">
&#x20; \<b>One application. Multiple APIs. Your daily dose of inspiration.\</b>
\</p>

---

## 📌 About The Project

**Daily Motivation Hub** is a Python-based command-line application that brings together multiple public APIs to deliver daily advice, inspirational quotes, book recommendations and random dog images.

The application provides an interactive terminal experience using Rich and Colorama, while storing retrieved information locally in Excel workbooks using OpenPyXL.

The goal of this project is to explore API integration, data retrieval, Python automation, file handling and interactive terminal application development.

---

## ✨ Features

| Feature                 | Description                                           |
| :---------------------- | :---------------------------------------------------- |
| 💡 Daily Advice         | Fetch random advice through a public API              |
| ✨ Inspirational Quotes | Retrieve and display random quotes                    |
| 📚 Book Recommendations | Discover books using Open Library                     |
| 🐶 Random Dog Images    | Fetch and download random dog images                  |
| 📊 Data Statistics      | View stored records across categories                 |
| ⚡ Get Everything       | Retrieve all four content categories in one operation |
| 📁 Excel Storage        | Save retrieved information in Excel workbooks         |
| 🎨 Interactive Terminal | Colorful terminal interface with progress animations  |

---

## 🛠️ Technology Stack

### Programming Language

\<p>
&#x20; \<img src="[https://skillicons.dev/icons?i=python](https://skillicons.dev/icons?i=python)" height="42" alt="Python"/>
\</p>

### Python Libraries

\<p>
&#x20; \<img src="[https://img.shields.io/badge/Requests-HTTP%20Client-2CA5E0?style=flat-square](<https://img.shields.io/badge/Requests-HTTP%20Client-2CA5E0?style=flat-square>)"/>
&#x20; \<img src="[https://img.shields.io/badge/Rich-Terminal%20UI-555555?style=flat-square](<https://img.shields.io/badge/Rich-Terminal%20UI-555555?style=flat-square>)"/>
&#x20; \<img src="[https://img.shields.io/badge/Colorama-Terminal%20Colors-008000?style=flat-square](<https://img.shields.io/badge/Colorama-Terminal%20Colors-008000?style=flat-square>)"/>
&#x20; \<img src="[https://img.shields.io/badge/OpenPyXL-Excel%20Automation-217346?style=flat-square](<https://img.shields.io/badge/OpenPyXL-Excel%20Automation-217346?style=flat-square>)"/>
\</p>

- **Requests:** Communicating with external REST APIs.
- **Rich:** Interactive terminal panels, tables and progress animations.
- **Colorama:** Colored terminal output.
- **OpenPyXL:** Creating, reading and updating Excel workbooks.
- **Pathlib:** Managing application directories.
- **Datetime:** Recording timestamps for stored information.

---

## 🌐 Public APIs Integrated

The application uses the following public API endpoints in its current implementation.

| API                     | Purpose              | Endpoint                                   |
| :---------------------- | :------------------- | :----------------------------------------- |
| Advice Slip / ZenQuotes | Daily advice         | [ZenQuotes](https://zenquotes.io/)          |
| Type.fit                | Inspirational quotes | [Type.fit API](https://type.fit/api/quotes) |
| Open Library            | Book recommendations | [Open Library](https://openlibrary.org/)    |
| Dog CEO                 | Random dog images    | [Dog CEO API](https://dog.ceo/dog-api/)     |

**API data flow:**

1. The application sends an HTTP request.
2. The API returns a JSON response.
3. Python processes the returned information.
4. Rich displays the information in the terminal.
5. OpenPyXL stores the information in Excel files.
6. Downloaded dog images are stored locally and embedded in the dog workbook.

---

## 🖥️ Application Menu

The application provides seven interactive options:

| Option | Function                    |
| :----: | :-------------------------- |
|   1   | Get Daily Advice            |
|   2   | Get Inspirational Quote     |
|   3   | Get Book Recommendation     |
|   4   | View Application Statistics |
|   5   | Get Random Dog Image        |
|   6   | Get All Data                |
|   7   | Exit Application            |

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
├── .gitignore
├── requirements.txt
└── README.md
```

**Directory details:**

- `data/` – Stores retrieved API information in Excel workbooks.
- `images/` – Stores downloaded dog images.
- `api-hub-final.py` – An alternative application implementation.
- `finale_main (1).py` – An earlier application implementation.
- `test.py` – The most complete-looking interactive application version in this repository.

---

## ⚙️ Installation & Setup

### 1. Clone the Repository

```bash
git clone https://github.com/umbaupesrehaan1-art/api-hub.git
```

### 2. Navigate to the Project

```bash
cd api-hub
```

### 3. Install Python

Install Python 3.10 or higher.

Check your Python version:

```bash
python --version
```

### 4. Install Dependencies

```bash
pip install requests rich colorama openpyxl
```

Alternatively, install from the provided `requirements.txt`:

```bash
pip install -r requirements.txt
```

### 5. Run the Application

```bash
python test.py
```

The application initializes its local data directories and Excel workbooks when needed.

---

## 📊 Data Storage

The application uses Excel workbooks for local data persistence.

| Workbook        | Stored Information                                          |
| :-------------- | :---------------------------------------------------------- |
| `advice.xlsx` | Advice ID, advice text and timestamp                        |
| `quotes.xlsx` | Quote, author and timestamp                                 |
| `books.xlsx`  | Book title, author, publication information and timestamp   |
| `dogs.xlsx`   | Image URL, breed information, timestamp and embedded images |

Downloaded images are stored in the `images/` directory.

---

## 🔍 What I Learned

Through this project, I explored and practised:

- Integrating multiple external REST APIs.
- Sending HTTP requests and processing JSON responses.
- Handling local files and directories using Python.
- Creating and managing Excel workbooks programmatically.
- Building interactive command-line applications.
- Working with terminal animations and progress indicators.
- Organizing application data and maintaining local records.
- Combining different Python libraries into one application.

---

## 🚧 Future Improvements

Potential improvements for future versions:

- [ ] Improve API error handling and response validation.
- [ ] Add a dedicated configuration module.
- [ ] Introduce structured logging.
- [ ] Improve automated testing.
- [ ] Add more public API integrations.
- [ ] Add an interactive graphical user interface.
- [ ] Add data export and reporting functionality.
- [ ] Improve application architecture and modularity.

---

## 👨‍💻 Developer

**Rehaan Usama Umba**

Computer Engineering Student | AI & Software Developer

Interested in Python, Artificial Intelligence, API integration and software engineering.

\<p>
&#x20; \<a href="[https://github.com/umbaupesrehaan1-art](https://github.com/umbaupesrehaan1-art)">
&#x20;   \<img src="[https://img.shields.io/badge/GitHub-Profile-181717?style=for-the-badge&amp;logo=github&amp;logoColor=white](https://img.shields.io/badge/GitHub-Profile-181717?style=for-the-badge\&logo=github\&logoColor=white)"/>
&#x20; \</a>
&#x20; \<a href="mailto\:umbaupesrehaan1\@gmail.com">
&#x20;   \<img src="[https://img.shields.io/badge/Email-Contact-EA4335?style=for-the-badge&amp;logo=gmail&amp;logoColor=white](https://img.shields.io/badge/Email-Contact-EA4335?style=for-the-badge\&logo=gmail\&logoColor=white)"/>
&#x20; \</a>
\</p>

**Email:** [umbaupesrehaan1@gmail.com](mailto\:umbaupesrehaan1@gmail.com)

**Location:** Mumbai, Maharashtra, India

---

\<p align="center">
&#x20; \<b>Learn. Build. Experiment. Improve.\</b>
&#x20; \<br/>
&#x20; ⭐ Thank you for visiting Daily Motivation Hub!
\</p>
