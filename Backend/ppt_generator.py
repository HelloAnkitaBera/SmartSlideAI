import os
import json
import google.generativeai as genai
from dotenv import load_dotenv

from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor


# =====================================
# LOAD ENV
# =====================================

load_dotenv()

api_key = os.getenv("GEMINI_API_KEY")

if not api_key:
    raise ValueError("GEMINI_API_KEY missing in .env")

genai.configure(api_key=api_key)

model = genai.GenerativeModel("gemini-1.5-flash")


# =====================================
# THEMES
# =====================================

THEMES = {
    "default": {
        "primary": RGBColor(30, 144, 255),
        "text": RGBColor(40, 40, 40),
    },

    "dark": {
        "primary": RGBColor(255, 255, 255),
        "text": RGBColor(220, 220, 220),
    }
}


# =====================================
# FALLBACK CONTENT
# =====================================

def fallback_content(topic, num_slides):

    slides = []

    for i in range(num_slides):

        slides.append({
            "title": f"{topic} - Section {i+1}",

            "content": [{
                "hook": f"Interesting fact about {topic}",

                "concept": f"Core Concept {i+1}",

                "explanation": (
                    f"This slide explains topic area {i+1} "
                    f"related to {topic}."
                ),

                "example": (
                    f"Real-world example for section {i+1}"
                ),

                "impact": (
                    f"Important impact of section {i+1}"
                )
            }]
        })

    return slides


# =====================================
# CLEAN TEXT
# =====================================

def clean_text(text):

    return str(text).encode(
        "utf-8",
        "ignore"
    ).decode("utf-8")


# =====================================
# GENERATE CONTENT
# =====================================

def generate_content(topic, num_slides=5):

    prompt = f"""
Create EXACTLY {num_slides} UNIQUE presentation slides on:

{topic}

Rules:
- Every slide MUST discuss a DIFFERENT subtopic
- Avoid repeating information
- Make explanations detailed
- Return ONLY valid JSON

Format:

[
  {{
    "title": "Slide Title",
    "content": [
      {{
        "hook": "Interesting opening",
        "concept": "Main concept",
        "explanation": "Detailed explanation",
        "example": "Real world example",
        "impact": "Importance or impact"
      }}
    ]
  }}
]

NO markdown.
ONLY JSON.
"""

    try:

        response = model.generate_content(prompt)

        text = response.text.strip()

        text = text.replace("```json", "")
        text = text.replace("```", "")
        text = text.strip()

        start = text.find("[")
        end = text.rfind("]") + 1

        text = text[start:end]

        slides = json.loads(text)

        if not isinstance(slides, list):
            return fallback_content(topic, num_slides)

        cleaned = []

        for slide in slides:

            cleaned_slide = {
                "title": clean_text(
                    slide.get("title", "Slide")
                ),

                "content": []
            }

            for item in slide.get("content", []):

                cleaned_slide["content"].append({

                    "hook": clean_text(
                        item.get("hook", "")
                    ),

                    "concept": clean_text(
                        item.get("concept", "")
                    ),

                    "explanation": clean_text(
                        item.get("explanation", "")
                    ),

                    "example": clean_text(
                        item.get("example", "")
                    ),

                    "impact": clean_text(
                        item.get("impact", "")
                    )
                })

            cleaned.append(cleaned_slide)

        while len(cleaned) < num_slides:

            cleaned.extend(
                fallback_content(
                    topic,
                    1
                )
            )

        return cleaned[:num_slides]

    except Exception as e:

        print("Gemini Error:", e)

        return fallback_content(
            topic,
            num_slides
        )


# =====================================
# PPT GENERATOR
# =====================================

class PPTGenerator:

    def __init__(self, theme="default"):

        self.prs = Presentation()

        self.theme = THEMES.get(
            theme,
            THEMES["default"]
        )

    # =================================
    # TITLE SLIDE
    # =================================

    def add_title_slide(self, topic):

        slide = self.prs.slides.add_slide(
            self.prs.slide_layouts[0]
        )

        title = slide.shapes.title

        title.text = clean_text(topic)

        p = title.text_frame.paragraphs[0]

        p.font.size = Pt(30)
        p.font.bold = True
        p.font.color.rgb = self.theme["primary"]

        try:

            subtitle = slide.placeholders[1]

            subtitle.text = (
                "AI Generated Presentation"
            )

            sp = subtitle.text_frame.paragraphs[0]

            sp.font.size = Pt(18)

        except:
            pass

    # =================================
    # CONTENT SLIDE
    # =================================

    def add_slide(self, title, content):

        slide = self.prs.slides.add_slide(
            self.prs.slide_layouts[1]
        )

        slide.shapes.title.text = clean_text(title)

        title_para = (
            slide.shapes.title
            .text_frame.paragraphs[0]
        )

        title_para.font.size = Pt(24)

        title_para.font.bold = True

        title_para.font.color.rgb = (
            self.theme["primary"]
        )

        body = slide.placeholders[1]

        tf = body.text_frame

        tf.clear()

        tf.word_wrap = True

        for item in content:

            sections = [

                f"Hook: {item['hook']}",

                f"Concept: {item['concept']}",

                f"Explanation: {item['explanation']}",

                f"Example: {item['example']}",

                f"Impact: {item['impact']}"
            ]

            for i, line in enumerate(sections):

                if i == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()

                p.text = clean_text(line)[:600]

                p.font.size = Pt(16)

                p.font.color.rgb = (
                    self.theme["text"]
                )

    # =================================
    # SAVE PPT
    # =================================

    def save(self, topic):

        output_dir = os.path.abspath(os.path.join(
            os.path.dirname(__file__),
            "..",
            "Frontend",
            "generated"
        ))

        os.makedirs(
            output_dir,
            exist_ok=True
        )

        safe_name = "".join(

            c for c in topic

            if c.isalnum() or c in (" ", "_")
        )

        filename = (
            safe_name.replace(" ", "_")
            + ".pptx"
        )

        path = os.path.join(
            output_dir,
            filename
        )

        if os.path.exists(path):
            os.remove(path)

        self.prs.save(path)

        # Validation
        Presentation(path)

        size = os.path.getsize(path)

        print(f"SUCCESSFULLY SAVED")
        print(f"SIZE: {size} bytes")
        print(f"LOCATION: {path}")

        return path

    # =================================
    # CREATE PPT
    # =================================

    def create_presentation(
        self,
        topic,
        num_slides=5
    ):

        self.add_title_slide(topic)

        slides = generate_content(
            topic,
            num_slides - 1
        )

        for slide in slides:

            self.add_slide(
                slide["title"],
                slide["content"]
            )

        return self.save(topic)


# =====================================
# MAIN FUNCTION
# =====================================

def create_ppt(
    topic,
    num_slides=5,
    theme="default"
):

    ppt = PPTGenerator(theme)

    return ppt.create_presentation(
        topic,
        num_slides
    )


# =====================================
# TEST
# =====================================

if __name__ == "__main__":

    try:

        ppt_path = create_ppt(
            topic="Cloud Computing Solutions",
            num_slides=5,
            theme="default"
        )

        print("\nPPT CREATED")
        print(ppt_path)

    except Exception as e:

        print("\nERROR:")
        print(e)