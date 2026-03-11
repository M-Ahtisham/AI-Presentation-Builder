# AI Powered Presentation Builder

This project is an **AI-powered presentation generator** that automatically creates PowerPoint slides from a user prompt.
It uses AI to generate structured content, fetches relevant images, and programmatically builds a professional presentation.

---

# Core Technologies at Play

### 1. AI (Gemini) API

Used for intelligent content generation and slide structuring.
The AI receives a topic prompt and generates structured presentation content in JSON format.

### 2. Image Generation (Unsplash) API

Provides high-quality images that are automatically added to slides to enhance visual presentation.

### 3. Python

Python orchestrates the entire system, handling:

* API communication
* Content processing
* Presentation generation
* Web interface logic

### 4. JSON

JSON is used as the structured data format for communication between components.
The AI outputs presentation content in JSON, which is then parsed to build slides.

### 5. python-pptx

The `python-pptx` library is used to programmatically create PowerPoint presentations (`.pptx`) from the generated content.

### 6. Streamlit

Streamlit provides a simple web interface where users can:

* Enter a presentation topic
* Generate slides
* Download the generated PowerPoint file

---

# Project Setup Guide

Follow the steps below to run the project locally.

---

## 1. Clone the Repository

```bash
git clone <https://github.com/M-Ahtisham/AI-Presentation-Builder.git>
cd <repository-folder>
```

---

## 2. Create a Virtual Environment

Create a virtual environment named `.venv`.

```bash
python3 -m venv .venv
```

---

## 3. Activate the Virtual Environment

### Linux / macOS

```bash
source .venv/bin/activate
```

### Windows (PowerShell)

```powershell
.venv\Scripts\Activate.ps1
```

After activation, your terminal should display `(.venv)` at the beginning of the prompt.

---

## 4. Install Dependencies

Install all required Python packages:

```bash
pip install -r requirements.txt
```

---

## 5. Configure API Keys

Create a `.env` file in the project root and add your API keys.

Example:

```
GEMINI_API_KEY=your_gemini_api_key
UNSPLASH_ACCESS_KEY=your_unsplash_key
```

---

## 6. Run the Application

Start the Streamlit application:

```bash
streamlit run app.py
```

This will launch the web interface in your browser.

---

# Example Workflow

1. User enters a **presentation topic** in the Streamlit interface.
2. The **Gemini API generates structured slide content**.
3. The system fetches **relevant images from Unsplash**.
4. `python-pptx` builds the **PowerPoint presentation**.
5. The user downloads the generated `.pptx` file.

---

# Updating Dependencies

If you install new libraries, update the requirements file:

```bash
pip freeze > requirements.txt
```

---

# Future Improvements

* Slide theme customization
* AI-generated diagrams and charts
* Multi-language presentation generation
* Export to PDF and Google Slides
