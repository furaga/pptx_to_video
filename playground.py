import os
import glob
from comtypes import client

PPT_NAME = "samples/from_pptx/sample.pptx"
OUT_DIR = "images"


def export_img(fname, odir):
    application = client.CreateObject("Powerpoint.Application")
    application.Visible = True
    current_folder = os.getcwd()

    presentation = application.Presentations.open(os.path.join(current_folder, fname))

    export_path = os.path.join(current_folder, odir)
    presentation.Export(export_path, FilterName="png")
    for slide in presentation.Slides:
        # Extract notes from slide
        notes = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
        notes = notes.replace("\r", "\n")

    presentation.close()
    application.quit()


def rename_img(odir):
    file_list = glob.glob(os.path.join(odir, "*.PNG"))
    for fname in file_list:
        new_fname = fname.replace("スライド", "slide").lower()
        os.rename(fname, new_fname)


if __name__ == "__main__":
    export_img(PPT_NAME, OUT_DIR)
    rename_img(OUT_DIR)
