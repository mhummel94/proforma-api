from fastapi import FastAPI
from pydantic import BaseModel, Field
from typing import List
from openpyxl import load_workbook
from io import BytesIO
from dotenv import load_dotenv
from pathlib import Path
import dropbox
import os

BASE_DIR = Path(__file__).resolve().parent
ENV_FILE = BASE_DIR / ".env"
load_dotenv(ENV_FILE)

app = FastAPI()


class SubjectProperty(BaseModel):
    address: str
    sqft: int | float | None = None
    beds: int | float | None = None
    baths: int | float | None = None
    year_built: int | None = None
    redfin_url: str | None = None


class Comp(BaseModel):
    address: str
    sqft: int | float | None = None
    beds: int | float | None = None
    baths: int | float | None = None
    year_built: int | None = None
    sold_date: str | None = None
    sold_price: int | float | None = None
    redfin_url: str | None = None


class ProformaRequest(BaseModel):
    dropbox_path: str
    subject_property: SubjectProperty
    comps: List[Comp] = Field(default_factory=list)


def get_dropbox_client():
    app_key = os.getenv("DROPBOX_APP_KEY")
    app_secret = os.getenv("DROPBOX_APP_SECRET")
    refresh_token = os.getenv("DROPBOX_REFRESH_TOKEN")

    if not app_key or not app_secret or not refresh_token:
        raise ValueError("Missing Dropbox credentials in .env")

    return dropbox.Dropbox(
        oauth2_refresh_token=refresh_token,
        app_key=app_key,
        app_secret=app_secret
    )


def format_bed_bath_year(beds, baths, year_built):
    bed_text = "" if beds is None else str(beds)
    bath_text = "" if baths is None else str(baths)
    year_text = "" if year_built is None else str(year_built)
    return f"{bed_text}/{bath_text}/{year_text}"


@app.get("/")
async def root():
    return {"message": "API is running"}


@app.get("/health")
async def health():
    return {
        "status": "ok",
        "dropbox_app_key_raw": os.getenv("DROPBOX_APP_KEY"),
        "dropbox_app_secret_raw": os.getenv("DROPBOX_APP_SECRET"),
        "dropbox_refresh_token_raw": os.getenv("DROPBOX_REFRESH_TOKEN"),
        "env_keys_with_dropbox": [k for k in os.environ.keys() if "DROPBOX" in k]
    }


@app.post("/populate-proforma")
async def populate_proforma(data: ProformaRequest):
    try:
        dbx = get_dropbox_client()
    except Exception as e:
        return {"success": False, "error": str(e)}

    # Download file from Dropbox
    try:
        metadata, response = dbx.files_download(data.dropbox_path)
        file_bytes = response.content
    except Exception as e:
        return {"success": False, "error": f"Download failed: {str(e)}"}

    try:
        workbook = load_workbook(BytesIO(file_bytes))
        sheet = workbook.active

        # Subject property
        sheet["B16"] = data.subject_property.address
        sheet["C16"] = data.subject_property.sqft
        sheet["D16"] = format_bed_bath_year(
            data.subject_property.beds,
            data.subject_property.baths,
            data.subject_property.year_built
        )
        sheet["B17"] = data.subject_property.redfin_url

        # Comps (rows 23–29)
        start_row = 23

        for i, comp in enumerate(data.comps[:7]):
            row = start_row + i
            sheet[f"B{row}"] = comp.address
            sheet[f"C{row}"] = comp.sqft
            sheet[f"D{row}"] = format_bed_bath_year(
                comp.beds,
                comp.baths,
                comp.year_built
            )
            sheet[f"E{row}"] = comp.sold_date
            sheet[f"F{row}"] = comp.sold_price
            sheet[f"H{row}"] = comp.redfin_url

        # Save to memory
        output = BytesIO()
        workbook.save(output)
        output.seek(0)

    except Exception as e:
        return {"success": False, "error": f"Processing failed: {str(e)}"}

    # Upload back to Dropbox
    try:
        dbx.files_upload(
            output.read(),
            data.dropbox_path,
            mode=dropbox.files.WriteMode.overwrite
        )
    except Exception as e:
        return {"success": False, "error": f"Upload failed: {str(e)}"}

    return {
        "success": True,
        "message": "Proforma populated successfully"
    }