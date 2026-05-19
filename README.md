# 📊 SmartSlide AI

SmartSlide AI is a professional, AI-powered presentation generator that transforms user-defined topics into beautifully structured, informative, and fully editable PowerPoint (`.pptx`) presentations in seconds. Leveraging the power of Google's Gemini API and the versatility of Python's `python-pptx` library, it builds content-rich slides featuring hooks, concepts, detailed explanations, real-world examples, and visual impacts.

---

## ✨ Features

- **Instant AI Content Generation**: Analyzes any topic and automatically structures your slides with original, high-quality content.
- **Custom Slide Counts**: Support for generating presentations of different lengths (5, 7, 10, or 15 slides).
- **Professional Themes**: Offers distinct presentation themes, such as **Corporate**, **Modern**, and **Startup**, dynamically styling fonts, titles, and text colors.
- **Rich Slide Structure**: Generates detailed content for each slide:
  - **Hook**: An engaging opening statement or interesting fact.
  - **Concept**: The core idea or focus of the slide.
  - **Explanation**: A deep dive into the subject matter.
  - **Example**: Real-world application or study.
  - **Impact**: The key takeaway or significance.
- **Premium Dark-Mode UI**: A modern, glassmorphic, and fully responsive web interface built with pure HTML/CSS/JS.
- **Fully Editable Outputs**: Download your presentation as standard Microsoft PowerPoint (`.pptx`) files, compatible with MS PowerPoint, Google Slides, and Keynote.

---

## 🛠️ Tech Stack

- **Frontend**: HTML5, CSS3 (Vanilla Custom Styles), JavaScript (ES6 Fetch API)
- **Backend Framework**: Flask (Python) with CORS support
- **AI Integration**: Google Generative AI (`gemini-1.5-flash`)
- **Presentation Engine**: `python-pptx`
- **Configuration & Setup**: `python-dotenv`, custom `setup.py` validation runner

---

## 📁 Project Structure

```text
SmartSlideAI/
│
├── Backend/
│   ├── .env                    # Environment variables (Gemini API Key)
│   ├── __init__.py
│   ├── generated/              # Output directory for generated PPTX files
│   └── ppt_generator.py        # Presentation generation engine & Gemini API caller
│
├── Frontend/
│   ├── app.py                  # Flask web server (serves UI & REST APIs)
│   ├── generated/              # Synced output folder for client downloads
│   ├── static/
│   │   ├── script.js           # Client-side form handlers & AJAX requests
│   │   └── style.css           # Custom CSS for the premium glassmorphic UI
│   └── templates/
│       └── index.html          # Main application page
│
├── requirements.txt            # Python dependencies
├── setup.py                    # Automatic dependency installation & validation script
└── README.md                   # Project documentation
```

---

## 🚀 Getting Started

### Prerequisites
- Python 3.8 or higher installed on your system.
- A **Google Gemini API Key** (Get one from [Google AI Studio](https://aistudio.google.com/)).

### 1. Set Up Environment Variables
Create a `.env` file inside the `Backend/` directory (or edit the existing one) and add your Google Gemini API key:
```env
GEMINI_API_KEY=your_gemini_api_key_here
```

### 2. Install Dependencies
Run the included `setup.py` script. It installs the required packages and verifies the environment:
```bash
python setup.py
```
This script handles the installation of:
- `python-pptx`
- `google-generativeai`
- `python-dotenv`
- `Flask`
- `flask-cors`

### 3. Run the Application
Start the Flask development server:
```bash
python Frontend/app.py
```
The server will start running on:
`http://localhost:5000`

### 4. Generate Presentations
1. Open your web browser and navigate to `http://localhost:5000`.
2. Input a presentation topic (e.g., *Digital Marketing Strategy*, *Sustainable Business Practices*).
3. Select your desired slide count and design theme.
4. Click **Generate Professional PPT**.
5. Once complete, the PPTX file will automatically download to your computer.

---

## 🔒 License
This project is licensed under the MIT License.
