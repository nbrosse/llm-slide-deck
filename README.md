# llm-slide-deck

This repository provides the source code featured in the blog post [Generating slide decks with an LLM: a comparison of formats](https://nbrosse.github.io/posts/llm-slides/llm-slides.html), which compares what a language model actually produces when asked to rebuild the same presentation in four different formats. The companion website is available at [LLM Slide Decks](https://nbrosse.github.io/llm-slide-deck/).

Starting from a real‑world presentation — [the EDF Producer Booklet (FET17CR)](livret_producteur_fet17cr_v1.1.pdf) — I asked Google Gemini 2.5 Pro to reproduce the first pages in several formats:

- Google Slides (Google Slides API)
- PowerPoint (python-pptx)
- HTML/CSS
- Quarto/Reveal.js

## A note on the generated code

The generated code here is preserved as the model wrote it, defects included. Those defects are the evidence the blog post is built on, so please do not read them as oversights, and do not fix them without replacing the corresponding discussion in the post.

The known ones, all in `powerpoint/create_powerpoint_slides.py`:

- The canvas is set to `Inches(10) x Inches(7.5)`, which is 4:3, directly beneath a comment announcing a widescreen 16:9 format. The exported PDF measures 720x540 points. The Quarto deck (1280x720) and the Google Slides export (720x405) are both correctly 16:9.
- `shape.fill.transparency` is assigned twice, but `python-pptx` has no `transparency` property on `FillFormat`. The assignment binds an attribute on the Python object, no alpha reaches the XML, and the title slide's intended semi-transparent overlay renders as opaque white.
- The title reads `LIVRET D'ACCEUIL`; the source booklet reads `ACCUEIL`. The other three formats spell it correctly.

None of these raises an error. The script runs to completion, exits zero, and writes a `.pptx` that opens — which is the point the post makes about generating slides through an API rather than as text.

The HTML/CSS deck in `raw/` contains no JavaScript and therefore has no slide navigation: it is six stacked `<section>` elements that you scroll through. That is also what the model produced, and it is also left as-is.

## Requirements

To install the requirements, use [uv](https://docs.astral.sh/uv/): 

```bash
uv sync
```

## Structure

The repository is organized as follows:

- `google/`: Google Slides API
- `powerpoint/`: PowerPoint (python-pptx)
- `raw/`: HTML/CSS
- `quarto/`: Quarto/Reveal.js
- `images/`: Images used in the slides
- `index.qmd`: The main page of the website
- `_quarto.yml`: The configuration file for the website
- `README.md`: This file

## Google Slides

In order to use the Google Slides API, you need to set up a project in the Google Cloud Console and enable the Google Slides API. Follow these steps:

1. Go to the [Google Cloud Console](https://console.cloud.google.com/).
2. Create a new project or select an existing project.
3. Enable the Google Slides API and the Google Drive API for your project.
4. Create credentials (OAuth 2.0 client IDs) for your application.
5. Download the credentials JSON file and save it as `credentials.json` in the google directory `google/`.

Then, execute the following command to create the slide deck on Google Slides:

```bash
uv run python google/create_slides.py
```

## Powerpoint

To create the slide deck on Powerpoint, execute the following command:

```bash
uv run python powerpoint/create_powerpoint_slides.py
```