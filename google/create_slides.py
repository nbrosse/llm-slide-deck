import os.path
from pathlib import Path
import pickle

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# --- Configuration ---
# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/presentations', 'https://www.googleapis.com/auth/drive']

# Asset URLs
EDF_LOGO_URL = 'https://i.imgur.com/kSg7T9g.png' # EDF Logo with transparent background
SLIDE1_BG_URL = 'https://i.imgur.com/k6lPqS8.jpeg'
SLIDE4_DIAGRAM_URL = 'https://i.imgur.com/t1m9rVv.png' # Pre-rendered diagram for simplicity

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

google_path = Path(__file__).parent

def get_slides_service():
    """Shows basic usage of the Slides API.
    Creates a new presentation.
    """
    creds = None
    if os.path.exists(google_path / 'token.pickle'):
        with open(google_path / 'token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                google_path / 'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open(google_path / 'token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('slides', 'v1', credentials=creds)
    return service

def create_presentation(service, title):
    body = {'title': title}
    presentation = service.presentations().create(body=body).execute()
    print(f"Created presentation with ID: {presentation.get('presentationId')}")
    return presentation

def execute_requests(service, presentation_id, requests):
    body = {'requests': requests}
    response = service.presentations().batchUpdate(
        presentationId=presentation_id, body=body).execute()
    return response

def create_slide_1(presentation_id):
    slide_id = "slide_1_id"
    requests = [
        # Set background image for the first slide
        {
            'updatePageProperties': {
                'objectId': presentation_id,
                'pageProperties': {
                    'pageBackgroundFill': {
                        'stretchedPictureFill': {
                            'contentUrl': SLIDE1_BG_URL
                        }
                    }
                }
            }
        },
        # Create the semi-transparent overlay
        {
            'createShape': {
                'objectId': 'slide1_overlay',
                'shapeType': 'RECTANGLE',
                'elementProperties': {
                    'pageObjectId': presentation_id,
                    'size': {'width': {'magnitude': 700, 'unit': 'PT'}, 'height': {'magnitude': 400, 'unit': 'PT'}},
                    'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 130, 'translateY': 70, 'unit': 'PT'}
                }
            }
        },
        {
            'updateShapeProperties': {
                'objectId': 'slide1_overlay',
                'shapeProperties': {
                    'shapeBackgroundFill': {
                        'solidFill': {
                            'color': {'rgbColor': WHITE_TEXT},
                            'alpha': 0.15
                        }
                    },
                    'outline': {
                        'outlineFill': {
                            'color': {'rgbColor': WHITE_TEXT},
                            'alpha': 0
                        }
                    }
                },
                'fields': 'shapeBackgroundFill,outline'
            }
        },
        # EDF Logo
        {
            'createImage': {
                'objectId': 'slide1_logo',
                'url': EDF_LOGO_URL,
                'elementProperties': {
                    'pageObjectId': presentation_id,
                    'size': {'width': {'magnitude': 100, 'unit': 'PT'}, 'height': {'magnitude': 40, 'unit': 'PT'}},
                    'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 160, 'translateY': 100, 'unit': 'PT'}
                }
            }
        },
        # Title
        {
            'createShape': {
                'objectId': 'slide1_title', 'shapeType': 'TEXT_BOX',
                'elementProperties': {
                    'pageObjectId': presentation_id,
                    'size': {'width': {'magnitude': 600, 'unit': 'PT'}, 'height': {'magnitude': 100, 'unit': 'PT'}},
                    'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 180, 'translateY': 150, 'unit': 'PT'}
                }
            }
        },
        {
            'insertText': {
                'objectId': 'slide1_title', 'text': "Appel d'offres\néolien terrestre", 'insertionIndex': 0
            }
        },
        {
            'updateTextStyle': {
                'objectId': 'slide1_title',
                'style': {
                    'fontSize': {'magnitude': 36, 'unit': 'PT'},
                    'foregroundColor': {'color': {'rgbColor': EDF_BLUE_DARK}},
                    'bold': True,
                    'fontFamily': 'Open Sans'
                },
                'fields': 'fontSize,foregroundColor,bold,fontFamily'
            }
        },
        # Subtitle
        {
            'createShape': {
                'objectId': 'slide1_subtitle', 'shapeType': 'TEXT_BOX',
                'elementProperties': {
                    'pageObjectId': presentation_id,
                    'size': {'width': {'magnitude': 600, 'unit': 'PT'}, 'height': {'magnitude': 50, 'unit': 'PT'}},
                    'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 180, 'translateY': 260, 'unit': 'PT'}
                }
            }
        },
        {'insertText': {'objectId': 'slide1_subtitle', 'text': "(publié par la Commission de Régulation de l'Energie le 28 Avril 2017)"}},
        {'updateTextStyle': {'objectId': 'slide1_subtitle', 'style': {'fontSize': {'magnitude': 12, 'unit': 'PT'}, 'foregroundColor': {'color': {'rgbColor': TEXT_GRAY}}, 'fontFamily': 'Open Sans'}, 'fields': '*'}},
        # Livret
        {
            'createShape': {
                'objectId': 'slide1_livret', 'shapeType': 'TEXT_BOX',
                'elementProperties': {
                    'pageObjectId': presentation_id,
                    'size': {'width': {'magnitude': 600, 'unit': 'PT'}, 'height': {'magnitude': 50, 'unit': 'PT'}},
                    'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 180, 'translateY': 350, 'unit': 'PT'}
                }
            }
        },
        {'insertText': {'objectId': 'slide1_livret', 'text': "LIVRET D'ACCUEIL PRODUCTEUR"}},
        {'updateTextStyle': {'objectId': 'slide1_livret', 'style': {'fontSize': {'magnitude': 24, 'unit': 'PT'}, 'foregroundColor': {'color': {'rgbColor': EDF_ORANGE}}, 'bold': True, 'fontFamily': 'Open Sans'}, 'fields': '*'}},
        # Divider line
        {
            'createLine': {
                'objectId': 'slide1_divider',
                'category': 'STRAIGHT',
                'elementProperties': {
                    'pageObjectId': presentation_id,
                    'size': {'width': {'magnitude': 640, 'unit': 'PT'}},
                    'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 160, 'translateY': 330, 'unit': 'PT'}
                }
            }
        },
        {
            'updateLineProperties': {
                'objectId': 'slide1_divider',
                'lineProperties': {
                    'lineFill': {
                        'solidFill': {
                            'color': {'rgbColor': TEXT_GRAY},
                            'alpha': 0.3
                        }
                    }
                },
                'fields': 'lineFill'
            }
        }
    ]
    return requests

def create_slide_2(presentation_id):
    slide_id = "slide_2_id"
    requests = [
        {'createSlide': {'objectId': slide_id, 'slideLayoutReference': {'predefinedLayout': 'BLANK'}}},
        # Title
        {'createShape': {'objectId': 's2_title', 'shapeType': 'TEXT_BOX', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': 400, 'unit': 'PT'}, 'height': {'magnitude': 50, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 50, 'translateY': 50, 'unit': 'PT'}}}},
        {'insertText': {'objectId': 's2_title', 'text': 'SOMMAIRE'}},
        {'updateTextStyle': {'objectId': 's2_title', 'style': {'fontSize': {'magnitude': 36, 'unit': 'PT'}, 'foregroundColor': TEXT_GRAY, 'fontFamily': 'Open Sans', 'bold': True}, 'fields': '*'}},
    ]
    
    items = [
        ("Préambule", {'rgbColor': {'red': 0.16, 'green': 0.47, 'blue': 1.0}}),
        ("Présentation des acteurs", {'rgbColor': {'red': 0.0, 'green': 0.18, 'blue': 0.38}}),
        ("Parcours de contractualisation", {'rgbColor': {'red': 0.33, 'green': 0.55, 'blue': 0.18}}),
        ("Check-list des démarches", {'rgbColor': {'red': 0.68, 'green': 0.71, 'blue': 0.17}}),
        ("Questions - Réponses", {'rgbColor': {'red': 0.96, 'green': 0.49, 'blue': 0.0}}),
        ("Adresses utiles", {'rgbColor': {'red': 0.9, 'green': 0.29, 'blue': 0.1}}),
    ]

    y_pos = 140
    for i, (text, color) in enumerate(items):
        shape_id = f"s2_item_shape_{i}"
        text_id = f"s2_item_text_{i}"
        requests.extend([
            {'createShape': {'objectId': shape_id, 'shapeType': 'RECTANGLE', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': 500, 'unit': 'PT'}, 'height': {'magnitude': 40, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 50, 'translateY': y_pos, 'unit': 'PT'}}}},
            {'updateShapeProperties': {'objectId': shape_id, 'shapeProperties': {'shapeBackgroundFill': {'solidFill': {'color': {'rgbColor': color}}}}, 'fields': 'shapeBackgroundFill'}},
            {'insertText': {'objectId': shape_id, 'text': text}},
            {'updateTextStyle': {'objectId': shape_id, 'style': {'fontSize': {'magnitude': 16, 'unit': 'PT'}, 'foregroundColor': {'color': {'rgbColor': WHITE_TEXT}}, 'bold': True, 'fontFamily': 'Open Sans'}, 'fields': '*'}},
            {'updateParagraphStyle': {'objectId': shape_id, 'style': {'alignment': 'START'}, 'fields': 'alignment'}}
        ])
        y_pos += 55

    return requests

def create_slide_3(presentation_id):
    slide_id = "slide_3_id"
    requests = [
        {'createSlide': {'objectId': slide_id, 'slideLayoutReference': {'predefinedLayout': 'BLANK'}}},
        # Header bar
        {'createShape': {'objectId': 's3_header', 'shapeType': 'RECTANGLE', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': PPT_WIDTH, 'unit': 'PT'}, 'height': {'magnitude': 80, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 0, 'translateY': 0, 'unit': 'PT'}}}},
        {'updateShapeProperties': {'objectId': 's3_header', 'shapeProperties': {'shapeBackgroundFill': {'solidFill': {'color': {'rgbColor': HEADER_ORANGE}}}, 'outline': {'outlineFill': {'color': {'rgbColor': HEADER_ORANGE}}}}, 'fields': 'shapeBackgroundFill,outline'}},
        # Header Title
        {'createShape': {'objectId': 's3_title', 'shapeType': 'TEXT_BOX', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': 400, 'unit': 'PT'}, 'height': {'magnitude': 50, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 50, 'translateY': 15, 'unit': 'PT'}}}},
        {'insertText': {'objectId': 's3_title', 'text': 'Préambule'}},
        {'updateTextStyle': {'objectId': 's3_title', 'style': {'fontSize': {'magnitude': 36, 'unit': 'PT'}, 'foregroundColor': {'color': {'rgbColor': WHITE_TEXT}}, 'fontFamily': 'Open Sans', 'bold': True}, 'fields': '*'}},
        # Left column text
        {'createShape': {'objectId': 's3_left_col', 'shapeType': 'TEXT_BOX', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': 500, 'unit': 'PT'}, 'height': {'magnitude': 400, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 50, 'translateY': 120, 'unit': 'PT'}}}},
        {'insertText': {'objectId': 's3_left_col', 'text': "Ce document s'adresse uniquement aux lauréats de l'appel d'offres...\nCe document résume, sous une forme simplifiée, les étapes nécessaires...\nDans le cadre des missions de service public..."}},
        {'updateTextStyle': {'objectId': 's3_left_col', 'style': {'fontSize': {'magnitude': 14, 'unit': 'PT'}, 'fontFamily': 'Open Sans'}, 'fields': 'fontSize,fontFamily'}},
        {'createParagraphBullets': {'objectId': 's3_left_col', 'bulletPreset': 'DIAMOND', 'textRange': {'type': 'ALL'}}},
        {'updateTextStyle': {'objectId': 's3_left_col', 'style': {'foregroundColor': {'color': {'rgbColor': EDF_ORANGE}}}, 'textRange': {'type': 'ALL'}, 'fields': 'foregroundColor'}},
        
        # Right Column Info Boxes
        {'createShape': {'objectId': 's3_info1_box', 'shapeType': 'RECTANGLE', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': 320, 'unit': 'PT'}, 'height': {'magnitude': 130, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 600, 'translateY': 120, 'unit': 'PT'}}}},
        {'updateShapeProperties': {'objectId': 's3_info1_box', 'shapeProperties': {'shapeBackgroundFill': {'solidFill': {'color': {'rgbColor': INFO_BOX_BLUE}}}}, 'fields': 'shapeBackgroundFill'}},
        {'insertText': {'objectId': 's3_info1_box', 'text': "i   Ce livret ne saurait engager la responsabilité d'EDF quant aux obligations du producteur..."}},
        {'updateTextStyle': {'objectId': 's3_info1_box', 'style': {'fontSize': {'magnitude': 12, 'unit': 'PT'}, 'foregroundColor': {'color': {'rgbColor': WHITE_TEXT}}}, 'textRange': {'type': 'ALL'}, 'fields': 'fontSize,foregroundColor'}},
        {'updateTextStyle': {'objectId': 's3_info1_box', 'style': {'fontSize': {'magnitude': 30, 'unit': 'PT'}, 'italic': True, 'fontFamily': 'Times New Roman'}, 'textRange': {'type': 'FIXED_RANGE', 'startIndex': 0, 'endIndex': 1}, 'fields': 'fontSize,italic,fontFamily'}},
        
        {'createShape': {'objectId': 's3_info2_box', 'shapeType': 'RECTANGLE', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': 320, 'unit': 'PT'}, 'height': {'magnitude': 150, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 600, 'translateY': 270, 'unit': 'PT'}}}},
        {'updateShapeProperties': {'objectId': 's3_info2_box', 'shapeProperties': {'shapeBackgroundFill': {'solidFill': {'color': {'rgbColor': INFO_BOX_BLUE}}}}, 'fields': 'shapeBackgroundFill'}},
        {'insertText': {'objectId': 's3_info2_box', 'text': "i   Le lauréat s'engage à mettre en service et à exploiter une installation en tous points conforme..."}},
        {'updateTextStyle': {'objectId': 's3_info2_box', 'style': {'fontSize': {'magnitude': 12, 'unit': 'PT'}, 'foregroundColor': {'color': {'rgbColor': WHITE_TEXT}}}, 'textRange': {'type': 'ALL'}, 'fields': 'fontSize,foregroundColor'}},
        {'updateTextStyle': {'objectId': 's3_info2_box', 'style': {'fontSize': {'magnitude': 30, 'unit': 'PT'}, 'italic': True, 'fontFamily': 'Times New Roman'}, 'textRange': {'type': 'FIXED_RANGE', 'startIndex': 0, 'endIndex': 1}, 'fields': 'fontSize,italic,fontFamily'}},
    ]
    return requests

def create_slide_4(presentation_id):
    slide_id = "slide_4_id"
    requests = [
        {'createSlide': {'objectId': slide_id, 'slideLayoutReference': {'predefinedLayout': 'BLANK'}}},
        # Header bar
        {'createShape': {'objectId': 's4_header', 'shapeType': 'RECTANGLE', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': PPT_WIDTH, 'unit': 'PT'}, 'height': {'magnitude': 80, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 0, 'translateY': 0, 'unit': 'PT'}}}},
        {'updateShapeProperties': {'objectId': 's4_header', 'shapeProperties': {'shapeBackgroundFill': {'solidFill': {'color': {'rgbColor': HEADER_ORANGE}}}, 'outline': {'outlineFill': {'color': {'rgbColor': HEADER_ORANGE}}}}, 'fields': 'shapeBackgroundFill,outline'}},
        # Header Title
        {'createShape': {'objectId': 's4_title', 'shapeType': 'TEXT_BOX', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': 600, 'unit': 'PT'}, 'height': {'magnitude': 50, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 50, 'translateY': 15, 'unit': 'PT'}}}},
        {'insertText': {'objectId': 's4_title', 'text': 'Présentation des acteurs'}},
        {'updateTextStyle': {'objectId': 's4_title', 'style': {'fontSize': {'magnitude': 36, 'unit': 'PT'}, 'foregroundColor': {'color': {'rgbColor': WHITE_TEXT}}, 'fontFamily': 'Open Sans', 'bold': True}, 'fields': '*'}},
        # Diagram Image
        {
            'createImage': {
                'objectId': 's4_diagram',
                'url': SLIDE4_DIAGRAM_URL,
                'elementProperties': {
                    'pageObjectId': slide_id,
                    'size': {'width': {'magnitude': 800, 'unit': 'PT'}, 'height': {'magnitude': 400, 'unit': 'PT'}},
                    'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 80, 'translateY': 110, 'unit': 'PT'}
                }
            }
        },
    ]
    return requests

def create_slide_5(presentation_id):
    slide_id = "slide_5_id"
    requests = [
        {'createSlide': {'objectId': slide_id, 'slideLayoutReference': {'predefinedLayout': 'BLANK'}}},
        # Header bar & Title
        {'createShape': {'objectId': 's5_header', 'shapeType': 'RECTANGLE', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': PPT_WIDTH, 'unit': 'PT'}, 'height': {'magnitude': 80, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 0, 'translateY': 0, 'unit': 'PT'}}}},
        {'updateShapeProperties': {'objectId': 's5_header', 'shapeProperties': {'shapeBackgroundFill': {'solidFill': {'color': {'rgbColor': HEADER_ORANGE}}}, 'outline': {'outlineFill': {'color': {'rgbColor': HEADER_ORANGE}}}}, 'fields': 'shapeBackgroundFill,outline'}},
        {'createShape': {'objectId': 's5_title', 'shapeType': 'TEXT_BOX', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': 600, 'unit': 'PT'}, 'height': {'magnitude': 50, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 50, 'translateY': 15, 'unit': 'PT'}}}},
        {'insertText': {'objectId': 's5_title', 'text': 'Parcours de contractualisation'}},
        {'updateTextStyle': {'objectId': 's5_title', 'style': {'fontSize': {'magnitude': 36, 'unit': 'PT'}, 'foregroundColor': WHITE_TEXT, 'fontFamily': 'Open Sans', 'bold': True}, 'fields': '*'}},
    ]
    
    steps = [
        ("Demande de raccordement", "J'effectue ma demande de raccordement auprès du gestionnaire de réseau (maximum 2 mois après la désignation).", {'rgbColor': {'red': 0.1, 'green': 0.46, 'blue': 0.82}}),
        ("Demande de contrat", "Au plus près de l'achèvement de mon installation, j'envoie ma demande de contrat à EDF OA...", {'rgbColor': {'red': 0.19, 'green': 0.25, 'blue': 0.62}}),
        ("Notification de la date...", "Je notifie à EDF OA la date projetée de prise d'effet de mon contrat...", {'rgbColor': {'red': 0.98, 'green': 0.75, 'blue': 0.17}}),
        ("Mise en service du raccordement", "Je prends rendez vous avec mon gestionnaire de réseau pour mettre en service...", {'rgbColor': {'red': 0.96, 'green': 0.49, 'blue': 0.0}}),
        ("Achèvement de l'installation...", "J'achève mon installation dans un délai de 36 mois...", {'rgbColor': {'red': 0.83, 'green': 0.18, 'blue': 0.18}}),
        ("Signature du contrat...", "Dans le cadre du processus de signature, EDF OA m'adresse mon contrat...", {'rgbColor': {'red': 0.22, 'green': 0.56, 'blue': 0.22}}),
        ("Facture et règlement", "Je mets mes factures mensuellement, sur la base des données de facturation...", {'rgbColor': {'red': 0.48, 'green': 0.12, 'blue': 0.64}})
    ]
    
    y_pos = 100
    row_height = 55
    for i, (step_text, detail_text, color) in enumerate(steps):
        step_box_id = f's5_step_{i}'
        detail_box_id = f's5_detail_{i}'
        
        # Step Box
        requests.extend([
            {'createShape': {'objectId': step_box_id, 'shapeType': 'RECTANGLE', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': 350, 'unit': 'PT'}, 'height': {'magnitude': row_height, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 50, 'translateY': y_pos, 'unit': 'PT'}}}},
            {'updateShapeProperties': {'objectId': step_box_id, 'shapeProperties': {'shapeBackgroundFill': {'solidFill': {'color': {'rgbColor': color}}}}, 'fields': 'shapeBackgroundFill'}},
            {'insertText': {'objectId': step_box_id, 'text': f"{i+1}   {step_text}"}},
            {'updateTextStyle': {'objectId': step_box_id, 'style': {'fontSize': {'magnitude': 12, 'unit': 'PT'}, 'foregroundColor': {'color': {'rgbColor': WHITE_TEXT}}, 'bold':True}, 'fields': 'fontSize,foregroundColor,bold'}},
            {'updateTextStyle': {'objectId': step_box_id, 'style': {'fontSize': {'magnitude': 24, 'unit': 'PT'}}, 'textRange': {'type': 'FIXED_RANGE', 'startIndex': 0, 'endIndex': 1}, 'fields': 'fontSize'}},
            {'updateParagraphStyle': {'objectId': step_box_id, 'style': {'alignment': 'START', 'spaceAbove': {'magnitude':15, 'unit':'PT'}}, 'fields': '*'}},
        ])
        
        # Detail Box
        requests.extend([
            {'createShape': {'objectId': detail_box_id, 'shapeType': 'RECTANGLE', 'elementProperties': {'pageObjectId': slide_id, 'size': {'width': {'magnitude': 500, 'unit': 'PT'}, 'height': {'magnitude': row_height, 'unit': 'PT'}}, 'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 405, 'translateY': y_pos, 'unit': 'PT'}}}},
            {'updateShapeProperties': {'objectId': detail_box_id, 'shapeProperties': {'shapeBackgroundFill': {'solidFill': {'color': {'rgbColor': {'red': 0.96, 'green': 0.96, 'blue': 0.96}}}}}, 'fields': 'shapeBackgroundFill'}},
            {'insertText': {'objectId': detail_box_id, 'text': detail_text}},
            {'updateTextStyle': {'objectId': detail_box_id, 'style': {'fontSize': {'magnitude': 10, 'unit': 'PT'}, 'foregroundColor': {'color': {'rgbColor': TEXT_GRAY}}}, 'fields': '*'}},
            {'updateParagraphStyle': {'objectId': detail_box_id, 'style': {'alignment': 'START', 'spaceAbove': {'magnitude':5, 'unit':'PT'}, 'spaceBelow': {'magnitude':5, 'unit':'PT'}}, 'fields': '*'}},
        ])

        y_pos += row_height + 4

    return requests


def main():
    service = get_slides_service()
    presentation = create_presentation(service, "EDF - Appel d'offres éolien terrestre")
    presentation_id = presentation.get('presentationId')
    
    # Get the ID of the default first slide to delete it later
    default_slide_id = presentation['slides'][0]['objectId']

    print("Generating Slide 1: Title...")
    requests = create_slide_1(presentation['slides'][0]['objectId']) # Use existing slide for background
    
    print("Generating Slide 2: Sommaire...")
    requests.extend(create_slide_2(presentation_id))

    print("Generating Slide 3: Préambule...")
    requests.extend(create_slide_3(presentation_id))

    print("Generating Slide 4: Présentation des acteurs...")
    requests.extend(create_slide_4(presentation_id))
    
    print("Generating Slide 5: Parcours de contractualisation...")
    requests.extend(create_slide_5(presentation_id))
    
    # Delete the original blank slide now that we've added new ones
    # Note: The first slide was used for the title, so we're not deleting it.
    # If we created a new slide for the title, we would delete default_slide_id.
    
    print("Sending all requests to Google Slides API...")
    execute_requests(service, presentation_id, requests)
    
    print("\n--- All Done! ---")
    print(f"You can view your presentation at: https://docs.google.com/presentation/d/{presentation_id}")

if __name__ == '__main__':
    main()