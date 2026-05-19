from flask import (
    Flask,
    render_template,
    request,
    jsonify,
    send_file
)

from flask_cors import CORS

import importlib
import importlib.util
import os
import sys
import time


# =========================
# IMPORT PPT GENERATOR
# =========================

def _import_create_ppt():

    candidates = [

        "backend.ppt_generator",

        "ppt_generator",
    ]

    for name in candidates:

        try:

            mod = importlib.import_module(name)

            return getattr(mod, "create_ppt")

        except Exception as e:

            print(f"Import failed for {name}: {e}")

            continue

    here = os.path.dirname(__file__)

    parents = [

        here,

        os.path.abspath(
            os.path.join(here, "..")
        ),

        os.path.abspath(
            os.path.join(here, "..", "..")
        )
    ]

    for parent in parents:

        candidate = os.path.join(

            parent,

            "backend",

            "ppt_generator.py"
        )

        if os.path.isfile(candidate):

            try:

                spec = importlib.util.spec_from_file_location(

                    "ppt_generator_local",

                    candidate
                )

                if spec is None or spec.loader is None:
                    continue

                mod = importlib.util.module_from_spec(spec)

                sys.modules[spec.name] = mod

                spec.loader.exec_module(mod)

                return getattr(mod, "create_ppt")

            except Exception as e:

                print(f"Dynamic import error: {e}")

                continue

    raise ImportError(
        "Could not import create_ppt"
    )


# =========================
# LOAD FUNCTION
# =========================

create_ppt = _import_create_ppt()


# =========================
# FLASK APP
# =========================

import os
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, 'templates')
STATIC_DIR = os.path.join(BASE_DIR, 'static')

app = Flask(__name__, template_folder=TEMPLATE_DIR, static_folder=STATIC_DIR)

CORS(app)


# =========================
# HOME PAGE
# =========================

@app.route("/")
def home():

    return render_template(
        "index.html"
    )


# =========================
# GENERATE PPT
# =========================

@app.route(
    "/generate",
    methods=["POST"]
)

def generate():

    try:

        data = request.get_json()

        if not data:

            return jsonify({

                "error": "No JSON data received"

            }), 400

        topic = data.get("topic")

        num_slides = int(

            data.get("num_slides", 5)

        )

        theme = data.get(

            "theme",

            "default"
        )

        if not topic:

            return jsonify({

                "error": "Topic required"

            }), 400

        print("\n===================")
        print("GENERATING PPT")
        print("===================")

        print(f"TOPIC: {topic}")
        print(f"SLIDES: {num_slides}")
        print(f"THEME: {theme}")

        # =====================
        # GENERATE PPT
        # =====================

        ppt_path = create_ppt(

            topic,

            num_slides,

            theme
        )

        # WAIT FOR FILE WRITE
        time.sleep(1)

        # =====================
        # VALIDATE FILE
        # =====================

        if not ppt_path:

            return jsonify({

                "error": "PPT path is empty"

            }), 500

        if not os.path.exists(ppt_path):

            return jsonify({

                "error": "PPT file not created"

            }), 500

        file_size = os.path.getsize(
            ppt_path
        )

        print(f"PPT PATH: {ppt_path}")

        print(f"PPT SIZE: {file_size} bytes")



        # =====================
        # SUCCESS
        # =====================

        return jsonify({

            "success": True,

            "message": (
                "PPT generated successfully"
            ),

            "file_name": os.path.basename(
                ppt_path
            ),

            "download_url":
                f"/download?file="
                f"{os.path.basename(ppt_path)}"
        })

    except Exception as e:

        print("\nERROR DURING GENERATION:")
        print(str(e))

        return jsonify({

            "error": str(e)

        }), 500


# =========================
# DOWNLOAD ROUTE
# =========================

@app.route("/download")
def download():

    try:

        file = request.args.get("file")

        if not file:

            return jsonify({
                "error": "No file specified"
            }), 400

        generated_dir = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "generated"
        )

        path = os.path.join(
            generated_dir,
            file
        )

        print(f"DOWNLOAD PATH: {path}")

        if not os.path.exists(path):

            return jsonify({
                "error": "File not found"
            }), 404

        size = os.path.getsize(path)

        print(f"DOWNLOAD SIZE: {size}")

        return send_file(
            path,
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "presentationml.presentation"
            ),
            as_attachment=True,
            download_name=file
        )

    except Exception as e:

        print("DOWNLOAD ERROR:", e)

        return jsonify({
            "error": str(e)
        }), 500



# =========================
# RUN APP
# =========================

if __name__ == "__main__":

    app.run(

        debug=True,

        host="0.0.0.0",

        port=5000
    )