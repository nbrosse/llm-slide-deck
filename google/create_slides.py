import os.path
from pathlib import Path
import pickle
import mimetypes
from typing import Optional

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError


google_path = Path(__file__).parent
repo_path = google_path.parent

# --- Configuration ---
# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/presentations', 'https://www.googleapis.com/auth/drive']

# The title of the presentation to be created/deleted.
PRESENTATION_TITLE = "EDF - Appel d'offres éolien terrestre"

# Asset paths
edf_logo_path = repo_path / 'images/edf-logo.png' # EDF Logo with transparent background
slide1_bg_path = repo_path / 'images/slide-1-background.jpg'
slide3_acteurs_path = repo_path / 'images/slide-3-acteurs.jpg'

# Color Palette (Google Slides API expects just rgbColor dicts)
EDF_ORANGE = {'red': 1.0, 'green': 0.4, 'blue': 0.0}
EDF_BLUE_DARK = {'red': 0.0, 'green': 0.12, 'blue': 0.4}
HEADER_ORANGE = {'red': 0.96, 'green': 0.49, 'blue': 0.0}
TEXT_GRAY = {'red': 0.33, 'green': 0.33, 'blue': 0.33}
INFO_BOX_BLUE = {'red': 0.17, 'green': 0.24, 'blue': 0.31}
WHITE_TEXT = {'red': 1.0, 'green': 1.0, 'blue': 1.0}

# Presentation dimensions (16:9 aspect ratio in points)
PPT_WIDTH = 960
PPT_HEIGHT = 540


def get_google_creds():
    """Authenticate and return Google API credentials, saving token if needed."""
    creds = None
    token_path = google_path / 'token.pickle'
    creds_path = google_path / 'credentials.json'
    if token_path.exists():
        with open(token_path, 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                creds_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, 'wb') as token:
            pickle.dump(creds, token)
    return creds

def get_slides_service():
    """Return an authenticated Google Slides service object."""
    creds = get_google_creds()
    return build('slides', 'v1', credentials=creds)

def get_drive_service():
    """Return an authenticated Google Drive service object."""
    creds = get_google_creds()
    return build('drive', 'v3', credentials=creds)

def find_or_upload_image_to_drive(drive_service, file_path: Path) -> Optional[str]:
    """
    Search for an image on Google Drive by name. If it exists, return its ID.
    If not, upload it and return the new ID.
    """
    if not file_path.exists():
        print(f"Warning: Image file not found, skipping: {file_path}")
        return None

    file_name = file_path.name
    
    # Search for the file first to avoid duplicates
    try:
        query = f"name='{file_name}' and trashed=false"
        response = drive_service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
        files = response.get('files', [])
        if files:
            print(f"Image '{file_name}' already exists on Drive. Using existing ID: {files[0]['id']}")
            return files[0]['id']
    except HttpError as error:
        print(f"An error occurred while searching for file {file_name}: {error}")
        return None

    # If not found, upload the file
    print(f"Uploading '{file_name}' to Google Drive...")
    mimetype, _ = mimetypes.guess_type(str(file_path))
    if not mimetype:
        print(f"Error: Could not determine mimetype for {file_path}")
        return None
    
    file_metadata = {'name': file_name}
    media = MediaFileUpload(str(file_path), mimetype=mimetype)
    
    try:
        file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        file_id = file.get('id')
        print(f"Successfully uploaded '{file_name}'. File ID: {file_id}")
        # Set permission so anyone with the link can view (needed for Slides API to access the image)
        try:
            drive_service.permissions().create(
                fileId=file_id,
                body={
                    'type': 'anyone',
                    'role': 'reader',
                },
                fields='id'
            ).execute()
            print(f"Set public read permission for '{file_name}' (ID: {file_id})")
        except HttpError as perm_error:
            print(f"Warning: Could not set public permission for {file_name}: {perm_error}")
        return file_id
    except HttpError as error:
        print(f"An error occurred during file upload for {file_name}: {error}")
        return None

def upload_all_images(drive_service) -> dict[str, str]:
    """Upload all required images to Google Drive and return their IDs."""
    uploaded_images = {}
    image_paths = {
        'edf_logo': edf_logo_path,
        'slide1_bg': slide1_bg_path,
        'slide3_acteurs': slide3_acteurs_path
    }
    
    for key, path in image_paths.items():
        image_id = find_or_upload_image_to_drive(drive_service, path)
        if image_id:
            uploaded_images[key] = image_id
            
    return uploaded_images

def find_and_delete_presentation_by_title(drive_service, title: str):
    """Finds and deletes a Google Slide presentation with a specific title."""
    try:
        # Escape single quotes in title for Drive API query
        safe_title = title.replace("'", "\\'")
        query = f"name='{safe_title}' and mimeType='application/vnd.google-apps.presentation' and trashed=false"
        response = drive_service.files().list(q=query, fields='files(id, name)').execute()
        files = response.get('files', [])
        
        if not files:
            print(f"No existing presentation with title '{title}' found.")
            return

        for file in files:
            file_id = file.get('id')
            print(f"Found existing presentation '{file.get('name')}' (ID: {file_id}). Deleting it...")
            drive_service.files().delete(fileId=file_id).execute()
            print(f"Presentation ID {file_id} deleted successfully.")

    except HttpError as error:
        print(f"An error occurred while trying to delete the presentation: {error}")

def create_presentation(service, title):
    """Creates a new presentation."""
    body = {'title': title}
    presentation = service.presentations().create(body=body).execute()
    print(f"Created new presentation with ID: {presentation.get('presentationId')}")
    return presentation

def execute_requests(service, presentation_id, requests):
    """Executes a batch of requests to update a presentation."""
    if not requests:
        print("No requests to execute.")
        return
    body = {'requests': requests}
    response = service.presentations().batchUpdate(presentationId=presentation_id, body=body).execute()
    return response

def create_slide_1(slide_id: str, uploaded_images: dict[str, str]):
    """Generates requests for the title slide (Slide 1)."""
    # FIX: The `objectId` for page property updates and the `pageObjectId` for element
    # creation must be the ID of the SLIDE, not the presentation.
    requests = []
    
    if uploaded_images.get('slide1_bg'):
        requests.append({
            'updatePageProperties': {
                'objectId': slide_id, # FIX: Was presentation_id
                'pageProperties': {
                    'pageBackgroundFill': {
                        'stretchedPictureFill': {
                            'contentUrl': f"https://drive.google.com/uc?id={uploaded_images['slide1_bg']}"
                        }
                    }
                },
                'fields': 'pageBackgroundFill'
            }
        })
    
    requests.append({
        'createShape': {
            'objectId': 'slide1_overlay', 'shapeType': 'RECTANGLE',
            'elementProperties': {
                'pageObjectId': slide_id, # FIX: Was presentation_id
                'size': {'width': {'magnitude': 700, 'unit': 'PT'}, 'height': {'magnitude': 400, 'unit': 'PT'}},
                'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 130, 'translateY': 70, 'unit': 'PT'}
            }
        }
    })
    
    requests.append({
        'updateShapeProperties': {
            'objectId': 'slide1_overlay',
            'shapeProperties': {
                'shapeBackgroundFill': {
                    'solidFill': {'color': {'rgbColor': WHITE_TEXT}, 'alpha': 0.15}},
                'outline': {'outlineFill': {'solidFill': {'color': {'rgbColor': WHITE_TEXT}}}}
            },
            'fields': 'shapeBackgroundFill,outline'
        }
    })
    
    if uploaded_images.get('edf_logo'):
        requests.append({
            'createImage': {
                'objectId': 'slide1_logo',
                'url': f"https://drive.google.com/uc?id={uploaded_images['edf_logo']}",
                'elementProperties': {
                    'pageObjectId': slide_id, # FIX: Was presentation_id
                    'size': {'width': {'magnitude': 100, 'unit': 'PT'}, 'height': {'magnitude': 40, 'unit': 'PT'}},
                    'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 160, 'translateY': 100, 'unit': 'PT'}
                }
            }
        })
    
    # Title, Subtitle, and other elements... (all fixed to use slide_id)
    requests.extend([
        {'createShape': {'objectId': 'slide1_title', 'shapeType': 'TEXT_BOX', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': 600, 'unit': 'PT'}, 'height': {'magnitude': 100, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 180, 'translateY': 150, 'unit': 'PT'}}}},
        {'insertText': {'objectId': 'slide1_title', 'text': "Appel d'offres\néolien terrestre", 'insertionIndex': 0}},
        {'updateTextStyle': {'objectId': 'slide1_title', 'style': {'fontSize': {'magnitude': 36, 'unit': 'PT'}, 'foregroundColor': {'opaqueColor': {'rgbColor': EDF_BLUE_DARK}}, 'bold': True, 'fontFamily': 'Open Sans'}, 'fields': 'fontSize,foregroundColor,bold,fontFamily'}},
        {'createShape': {'objectId': 'slide1_subtitle', 'shapeType': 'TEXT_BOX', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': 600, 'unit': 'PT'}, 'height': {'magnitude': 50, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 180, 'translateY': 260, 'unit': 'PT'}}}},
        {'insertText': {'objectId': 'slide1_subtitle', 'text': "(publié par la Commission de Régulation de l'Energie le 28 Avril 2017)"}},
        {'updateTextStyle': {'objectId': 'slide1_subtitle', 'style': {'fontSize': {'magnitude': 12, 'unit': 'PT'}, 'foregroundColor': {'opaqueColor': {'rgbColor': TEXT_GRAY}}, 'fontFamily': 'Open Sans'}, 'fields': '*'}},
        {'createShape': {'objectId': 'slide1_livret', 'shapeType': 'TEXT_BOX', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': 600, 'unit': 'PT'}, 'height': {'magnitude': 50, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 180, 'translateY': 350, 'unit': 'PT'}}}},
        {'insertText': {'objectId': 'slide1_livret', 'text': "LIVRET D'ACCUEIL PRODUCTEUR"}},
        {'updateTextStyle': {'objectId': 'slide1_livret', 'style': {'fontSize': {'magnitude': 24, 'unit': 'PT'}, 'foregroundColor': {'opaqueColor': {'rgbColor': EDF_ORANGE}}, 'bold': True, 'fontFamily': 'Open Sans'}, 'fields': '*'}},
        {'createLine': {'objectId': 'slide1_divider', 'category': 'STRAIGHT', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': 640, 'unit': 'PT'}, 'height': {'magnitude': 0, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 160, 'translateY': 330, 'unit': 'PT'}}}},
        {'updateLineProperties': {'objectId': 'slide1_divider', 'lineProperties': {'lineFill': {'solidFill': {'color': {'rgbColor': TEXT_GRAY}, 'alpha': 0.3}}}, 'fields': 'lineFill'}}
    ])
    
    return requests

# Note: The following functions (create_slide_2, 3, 4, 5) were already correctly
# using the `slide_id` passed to them, so they didn't have the same critical bug.
# Minor improvements have been made where noted.

def create_slide_2(uploaded_images: dict[str, str]):
    slide_id = "slide_2_id"
    requests = [
        {'createSlide': {'objectId': slide_id, 'slideLayoutReference': {'predefinedLayout': 'BLANK'}}},
        {'createShape': {'objectId': 's2_title', 'shapeType': 'TEXT_BOX', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': 400, 'unit': 'PT'}, 'height': {'magnitude': 50, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 50, 'translateY': 50, 'unit': 'PT'}}}},
        {'insertText': {'objectId': 's2_title', 'text': 'SOMMAIRE'}},
        {'updateTextStyle': {'objectId': 's2_title', 'style': {'fontSize': {'magnitude': 36, 'unit': 'PT'}, 'foregroundColor': {'opaqueColor': {'rgbColor': TEXT_GRAY}}, 'fontFamily': 'Open Sans', 'bold': True}, 'fields': '*'}},
    ]
    items = [
        ("Préambule", {'red': 0.16, 'green': 0.47, 'blue': 1.0}),
        ("Présentation des acteurs", {'red': 0.0, 'green': 0.18, 'blue': 0.38}),
        ("Parcours de contractualisation", {'red': 0.33, 'green': 0.55, 'blue': 0.18}),
        ("Check-list des démarches", {'red': 0.68, 'green': 0.71, 'blue': 0.17}),
        ("Questions - Réponses", {'red': 0.96, 'green': 0.49, 'blue': 0.0}),
        ("Adresses utiles", {'red': 0.9, 'green': 0.29, 'blue': 0.1}),
    ]
    y_pos = 140
    for i, (text, color) in enumerate(items):
        shape_id = f"s2_item_shape_{i}"
        requests.extend([
            {'createShape': {'objectId': shape_id, 'shapeType': 'RECTANGLE', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': 500, 'unit': 'PT'}, 'height': {'magnitude': 40, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 50, 'translateY': y_pos, 'unit': 'PT'}}}},
            {'updateShapeProperties': {'objectId': shape_id, 'shapeProperties': {'shapeBackgroundFill': {'solidFill': {'color': {'rgbColor': color}}}}, 'fields': 'shapeBackgroundFill'}},
            {'insertText': {'objectId': shape_id, 'text': text}},
            {'updateTextStyle': {'objectId': shape_id, 'style': {'fontSize': {'magnitude': 16, 'unit': 'PT'}, 'foregroundColor': {'opaqueColor': {'rgbColor': WHITE_TEXT}}, 'bold': True, 'fontFamily': 'Open Sans'}, 'fields': '*'}},
            {'updateParagraphStyle': {'objectId': shape_id, 'style': {'alignment': 'START', 'indentStart': {'magnitude': 15, 'unit': 'PT'}, 'spaceAbove': {'magnitude': 8, 'unit': 'PT'}}, 'fields': 'alignment,indentStart,spaceAbove'}}
        ])
        y_pos += 55
    return requests

def create_slide_3(uploaded_images: dict[str, str]):
    slide_id = "slide_3_id"
    requests = [
        {'createSlide': {'objectId': slide_id, 'slideLayoutReference': {'predefinedLayout': 'BLANK'}}},
        {'createShape': {'objectId': 's3_header', 'shapeType': 'RECTANGLE', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': PPT_WIDTH, 'unit': 'PT'}, 'height': {'magnitude': 80, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 0, 'translateY': 0, 'unit': 'PT'}}}},
        {'updateShapeProperties': {'objectId': 's3_header', 'shapeProperties': {'shapeBackgroundFill': {'solidFill': {'color': {'rgbColor': HEADER_ORANGE}}}, 'outline': {'outlineFill': {'solidFill': {'color': {'rgbColor': HEADER_ORANGE}}}}}, 'fields': 'shapeBackgroundFill,outline'}},
        {'createShape': {'objectId': 's3_title', 'shapeType': 'TEXT_BOX', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': 400, 'unit': 'PT'}, 'height': {'magnitude': 50, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 50, 'translateY': 15, 'unit': 'PT'}}}},
        {'insertText': {'objectId': 's3_title', 'text': 'Préambule'}},
        {'updateTextStyle': {'objectId': 's3_title', 'style': {'fontSize': {'magnitude': 36, 'unit': 'PT'}, 'foregroundColor': {'opaqueColor': {'rgbColor': WHITE_TEXT}}, 'fontFamily': 'Open Sans', 'bold': True}, 'fields': '*'}},
        {'createShape': {'objectId': 's3_left_col', 'shapeType': 'TEXT_BOX', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': 500, 'unit': 'PT'}, 'height': {'magnitude': 400, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 50, 'translateY': 120, 'unit': 'PT'}}}},
        {'insertText': {'objectId': 's3_left_col', 'text': "Ce document s'adresse uniquement aux lauréats de l'appel d'offres...\nCe document résume, sous une forme simplifiée, les étapes nécessaires...\nDans le cadre des missions de service public..."}},
        # IMPROVEMENT: Set the text color to gray for readability before creating bullets.
        {'updateTextStyle': {'objectId': 's3_left_col', 'style': {'fontSize': {'magnitude': 14, 'unit': 'PT'}, 'fontFamily': 'Open Sans', 'foregroundColor': {'opaqueColor': {'rgbColor': TEXT_GRAY}}}, 'fields': 'fontSize,fontFamily,foregroundColor'}},
        # Use a valid bulletPreset value for Google Slides API
        {'createParagraphBullets': {'objectId': 's3_left_col', 'bulletPreset': 'BULLET_DISC_CIRCLE_SQUARE', 'textRange': {'type': 'ALL'}}},
        {'createShape': {'objectId': 's3_info1_box', 'shapeType': 'RECTANGLE', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': 320, 'unit': 'PT'}, 'height': {'magnitude': 130, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 600, 'translateY': 120, 'unit': 'PT'}}}},
        {'updateShapeProperties': {'objectId': 's3_info1_box', 'shapeProperties': {'shapeBackgroundFill': {'solidFill': {'color': {'rgbColor': INFO_BOX_BLUE}}}}, 'fields': 'shapeBackgroundFill'}},
        {'insertText': {'objectId': 's3_info1_box', 'text': "i   Ce livret ne saurait engager la responsabilité d'EDF quant aux obligations du producteur..."}},
        {'updateTextStyle': {'objectId': 's3_info1_box', 'style': {'fontSize': {'magnitude': 12, 'unit': 'PT'}, 'foregroundColor': {'opaqueColor': {'rgbColor': WHITE_TEXT}}}, 'textRange': {'type': 'ALL'}, 'fields': 'fontSize,foregroundColor'}},
        {'updateTextStyle': {'objectId': 's3_info1_box', 'style': {'fontSize': {'magnitude': 30, 'unit': 'PT'}, 'italic': True, 'fontFamily': 'Times New Roman'}, 'textRange': {'type': 'FIXED_RANGE', 'startIndex': 0, 'endIndex': 1}, 'fields': 'fontSize,italic,fontFamily'}},
        {'createShape': {'objectId': 's3_info2_box', 'shapeType': 'RECTANGLE', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': 320, 'unit': 'PT'}, 'height': {'magnitude': 150, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 600, 'translateY': 270, 'unit': 'PT'}}}},
        {'updateShapeProperties': {'objectId': 's3_info2_box', 'shapeProperties': {'shapeBackgroundFill': {'solidFill': {'color': {'rgbColor': INFO_BOX_BLUE}}}}, 'fields': 'shapeBackgroundFill'}},
        {'insertText': {'objectId': 's3_info2_box', 'text': "i   Le lauréat s'engage à mettre en service et à exploiter une installation en tous points conforme..."}},
        {'updateTextStyle': {'objectId': 's3_info2_box', 'style': {'fontSize': {'magnitude': 12, 'unit': 'PT'}, 'foregroundColor': {'opaqueColor': {'rgbColor': WHITE_TEXT}}}, 'textRange': {'type': 'ALL'}, 'fields': 'fontSize,foregroundColor'}},
        {'updateTextStyle': {'objectId': 's3_info2_box', 'style': {'fontSize': {'magnitude': 30, 'unit': 'PT'}, 'italic': True, 'fontFamily': 'Times New Roman'}, 'textRange': {'type': 'FIXED_RANGE', 'startIndex': 0, 'endIndex': 1}, 'fields': 'fontSize,italic,fontFamily'}},
    ]
    return requests

def create_slide_4(uploaded_images: dict[str, str]):
    slide_id = "slide_4_id"
    requests = [
        {'createSlide': {'objectId': slide_id, 'slideLayoutReference': {'predefinedLayout': 'BLANK'}}},
        {'createShape': {'objectId': 's4_header', 'shapeType': 'RECTANGLE', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': PPT_WIDTH, 'unit': 'PT'}, 'height': {'magnitude': 80, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 0, 'translateY': 0, 'unit': 'PT'}}}},
        {'updateShapeProperties': {'objectId': 's4_header', 'shapeProperties': {'shapeBackgroundFill': {'solidFill': {'color': {'rgbColor': HEADER_ORANGE}}}, 'outline': {'outlineFill': {'solidFill': {'color': {'rgbColor': HEADER_ORANGE}}}}}, 'fields': 'shapeBackgroundFill,outline'}},
        {'createShape': {'objectId': 's4_title', 'shapeType': 'TEXT_BOX', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': 600, 'unit': 'PT'}, 'height': {'magnitude': 50, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 50, 'translateY': 15, 'unit': 'PT'}}}},
        {'insertText': {'objectId': 's4_title', 'text': 'Présentation des acteurs'}},
        {'updateTextStyle': {'objectId': 's4_title', 'style': {'fontSize': {'magnitude': 36, 'unit': 'PT'}, 'foregroundColor': {'opaqueColor': {'rgbColor': WHITE_TEXT}}, 'fontFamily': 'Open Sans', 'bold': True}, 'fields': '*'}},
    ]
    if uploaded_images.get('slide3_acteurs'):
        requests.append({
            'createImage': {
                'objectId': 's4_diagram',
                'url': f"https://drive.google.com/uc?id={uploaded_images['slide3_acteurs']}",
                'elementProperties': {
                    'pageObjectId': slide_id,
                    'size': {'width': {'magnitude': 800, 'unit': 'PT'}, 'height': {'magnitude': 400, 'unit': 'PT'}},
                    'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 80, 'translateY': 110, 'unit': 'PT'}
                }
            }
        })
    return requests

def create_slide_5(uploaded_images: dict[str, str]):
    slide_id = "slide_5_id"
    requests = [
        {'createSlide': {'objectId': slide_id, 'slideLayoutReference': {'predefinedLayout': 'BLANK'}}},
        {'createShape': {'objectId': 's5_header', 'shapeType': 'RECTANGLE', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': PPT_WIDTH, 'unit': 'PT'}, 'height': {'magnitude': 80, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 0, 'translateY': 0, 'unit': 'PT'}}}},
        {'updateShapeProperties': {'objectId': 's5_header', 'shapeProperties': {'shapeBackgroundFill': {'solidFill': {'color': {'rgbColor': HEADER_ORANGE}}}, 'outline': {'outlineFill': {'solidFill': {'color': {'rgbColor': HEADER_ORANGE}}}}}, 'fields': 'shapeBackgroundFill,outline'}},
        {'createShape': {'objectId': 's5_title', 'shapeType': 'TEXT_BOX', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': 600, 'unit': 'PT'}, 'height': {'magnitude': 50, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 50, 'translateY': 15, 'unit': 'PT'}}}},
        {'insertText': {'objectId': 's5_title', 'text': 'Parcours de contractualisation'}},
        {'updateTextStyle': {'objectId': 's5_title', 'style': {'fontSize': {'magnitude': 36, 'unit': 'PT'}, 'foregroundColor': {'opaqueColor': {'rgbColor': WHITE_TEXT}}, 'fontFamily': 'Open Sans', 'bold': True}, 'fields': '*'}},
    ]
    steps = [
        ("Demande de raccordement", "J'effectue ma demande de raccordement auprès du gestionnaire de réseau (maximum 2 mois après la désignation).", {'red': 0.1, 'green': 0.46, 'blue': 0.82}),
        ("Demande de contrat", "Au plus près de l'achèvement de mon installation, j'envoie ma demande de contrat à EDF OA...", {'red': 0.19, 'green': 0.25, 'blue': 0.62}),
        ("Notification de la date...", "Je notifie à EDF OA la date projetée de prise d'effet de mon contrat...", {'red': 0.98, 'green': 0.75, 'blue': 0.17}),
        ("Mise en service du raccordement", "Je prends rendez vous avec mon gestionnaire de réseau pour mettre en service...", {'red': 0.96, 'green': 0.49, 'blue': 0.0}),
        ("Achèvement de l'installation...", "J'achève mon installation dans un délai de 36 mois...", {'red': 0.83, 'green': 0.18, 'blue': 0.18}),
        ("Signature du contrat...", "Dans le cadre du processus de signature, EDF OA m'adresse mon contrat...", {'red': 0.22, 'green': 0.56, 'blue': 0.22}),
        ("Facture et règlement", "Je mets mes factures mensuellement, sur la base des données de facturation...", {'red': 0.48, 'green': 0.12, 'blue': 0.64})
    ]
    y_pos = 100
    row_height = 55
    for i, (step_text, detail_text, color) in enumerate(steps):
        step_box_id = f's5_step_{i}'
        detail_box_id = f's5_detail_{i}'
        step_number_str = str(i + 1)
        # FIX: The end index for the number's style must be dynamic.
        step_number_end_index = len(step_number_str)

        requests.extend([
            {'createShape': {'objectId': step_box_id, 'shapeType': 'RECTANGLE', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': 350, 'unit': 'PT'}, 'height': {'magnitude': row_height, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 50, 'translateY': y_pos, 'unit': 'PT'}}}},
            {'updateShapeProperties': {'objectId': step_box_id, 'shapeProperties': {'shapeBackgroundFill': {'solidFill': {'color': {'rgbColor': color}}}}, 'fields': 'shapeBackgroundFill'}},
            {'insertText': {'objectId': step_box_id, 'text': f"{step_number_str}   {step_text}"}},
            {'updateTextStyle': {'objectId': step_box_id, 'style': {'fontSize': {'magnitude': 12, 'unit': 'PT'}, 'foregroundColor': {'opaqueColor': {'rgbColor': WHITE_TEXT}}, 'bold':True}, 'fields': 'fontSize,foregroundColor,bold'}},
            {'updateTextStyle': {'objectId': step_box_id, 'style': {'fontSize': {'magnitude': 24, 'unit': 'PT'}}, 'textRange': {'type': 'FIXED_RANGE', 'startIndex': 0, 'endIndex': step_number_end_index}, 'fields': 'fontSize'}}, # FIX
            {'updateParagraphStyle': {'objectId': step_box_id, 'style': {'alignment': 'START', 'spaceAbove': {'magnitude':15, 'unit':'PT'}, 'indentStart': {'magnitude': 15, 'unit': 'PT'}}, 'fields': '*'}},
            {'createShape': {'objectId': detail_box_id, 'shapeType': 'RECTANGLE', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': 500, 'unit': 'PT'}, 'height': {'magnitude': row_height, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 405, 'translateY': y_pos, 'unit': 'PT'}}}},
            {'updateShapeProperties': {'objectId': detail_box_id, 'shapeProperties': {'shapeBackgroundFill': {'solidFill': {'color': {'rgbColor': {'red': 0.96, 'green': 0.96, 'blue': 0.96}}}}}, 'fields': 'shapeBackgroundFill'}},
            {'insertText': {'objectId': detail_box_id, 'text': detail_text}},
            {'updateTextStyle': {'objectId': detail_box_id, 'style': {'fontSize': {'magnitude': 10, 'unit': 'PT'}, 'foregroundColor': {'opaqueColor': {'rgbColor': TEXT_GRAY}}}, 'fields': '*'}},
            {'updateParagraphStyle': {'objectId': detail_box_id, 'style': {'alignment': 'START', 'spaceAbove': {'magnitude':5, 'unit':'PT'}, 'spaceBelow': {'magnitude':5, 'unit':'PT'}, 'indentStart': {'magnitude': 10, 'unit': 'PT'}, 'indentEnd': {'magnitude': 10, 'unit': 'PT'}}, 'fields': '*'}},
        ])
        y_pos += row_height + 4
    return requests


def main():
    """Main function to generate the presentation."""
    # Get authenticated services
    slides_service = get_slides_service()
    drive_service = get_drive_service()
    
    # Clean up old presentation if it exists
    find_and_delete_presentation_by_title(drive_service, PRESENTATION_TITLE)

    # Find or upload all required images to Google Drive
    uploaded_images = upload_all_images(drive_service)
    
    # Create a new blank presentation
    presentation = create_presentation(slides_service, PRESENTATION_TITLE)
    presentation_id = presentation.get('presentationId')
    
    # Get the ID of the default first slide, which we will use for our title slide
    default_slide_id = presentation['slides'][0]['objectId']

    all_requests = []

    print("Generating requests for Slide 1: Title...")
    all_requests.extend(create_slide_1(default_slide_id, uploaded_images))
    
    print("Generating requests for Slide 2: Sommaire...")
    all_requests.extend(create_slide_2(uploaded_images))

    print("Generating requests for Slide 3: Préambule...")
    all_requests.extend(create_slide_3(uploaded_images))

    print("Generating requests for Slide 4: Présentation des acteurs...")
    all_requests.extend(create_slide_4(uploaded_images))
    
    print("Generating requests for Slide 5: Parcours de contractualisation...")
    all_requests.extend(create_slide_5(uploaded_images))
    
    # We are re-using the initial blank slide, so we don't need to delete it.
    # If we created a new title slide from scratch, we would add a deletion request:
    # all_requests.append({'deleteObject': {'objectId': default_slide_id}})
    
    print("\nSending all requests to Google Slides API in a single batch...")
    execute_requests(slides_service, presentation_id, all_requests)
    
    print("\n--- All Done! ---")
    print(f"You can view your presentation at: https://docs.google.com/presentation/d/{presentation_id}")

if __name__ == '__main__':
    main()