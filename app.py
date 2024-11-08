import os
import logging
from flask import Flask, redirect, render_template, request, send_from_directory, url_for

# Sets logging level to DEBUG.
logging.basicConfig(filename="record.log", level=logging.DEBUG)

# Start and configure the Flask app.
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 1024 * 1024 * 10  # 10 MB limit
app.config["UPLOAD_EXTENSIONS"] = ["xlsx", "xlsm"]
app.config["UPLOAD_PATH"] = "uploads"
app.config["PREFERRED_URL_SCHEME"] = "https"


# Filesize validation. Automatically detected by Flask based on the configuration.
@app.errorhandler(413)
def too_large(e):
    return "File is too large", 413


def allowed_file(filename):
    return (
        "." in filename
        and filename.rsplit(".", 1)[1].lower() in app.config["UPLOAD_EXTENSIONS"]
    )


# App Page Routing.
@app.route("/")
def index():
    app.logger.debug("Request for index page received")
    return render_template("index.html")


@app.route("/favicon.ico")
def favicon():
    app.logger.debug("Request for favicon received")
    return send_from_directory(
        os.path.join(app.root_path, "static"),
        "favicon.ico",
        mimetype="image/vnd.microsoft.icon",
    )


@app.route("/template_download", methods=["GET", "POST"])
def template_download():

    app.logger.debug("Request for template download received")
    return send_from_directory(
        os.path.join(app.root_path, "static"),
        "template.xlsm",  # Replace with your actual file name
        mimetype="application/vnd.ms-excel.sheet.macroEnabled.12",
        as_attachment=True
    )


@app.route("/input_upload", methods=["POST"])
def input_upload():
    app.logger.debug("Request for input upload received")
    if 'input_file' not in request.files:
        app.logger.debug("No file part in the request")
        return redirect(url_for('index'))

    input_file = request.files["input_file"]
    if input_file.filename == "":
        app.logger.debug("No file selected")
        return redirect(url_for('index'))

    if allowed_file(input_file.filename):
        # Gets the username from the email address in header.
        name = request.headers.get("X-MS-CLIENT-PRINCIPAL-NAME", "default_user").split("@")[0]
        # Creates a folder for the user if it doesn't exist.
        user_folder = os.path.join(app.config["UPLOAD_PATH"], name)
        os.makedirs(user_folder, exist_ok=True)

        # Saves the file to the user's folder. Always overwrites prior input.
        save_path = os.path.join(user_folder, "input.xlsx")
        input_file.save(save_path)
        
        app.logger.debug(f"User {name} input saved at {save_path}")
        return redirect(url_for('index'))
    else:
        app.logger.debug("File type not allowed")
        return redirect(url_for('index'))


@app.route("/output_download", methods=["POST"])
def output_download():
    app.logger.debug("Request for output download received")
    pass


if __name__ == "__main__":
    app.run(debug=True)
