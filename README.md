# llm-slide-deck

This repository provides the source code featured in the blog post [LLM-Powered Slide Decks: A Comparison of Formats](https://nbrosse.github.io/posts/llm-slides/llm-slides.html), which explores how large language models can streamline and enhance the process of creating slide presentations in various formats. The companion website is available at [LLM Slide Decks](https://nbrosse.github.io/llm-slide-deck/).

Starting from a real‑world presentation — [the EDF Producer Booklet (FET17CR)](livret_producteur_fet17cr_v1.1.pdf) — I asked Google Gemini 2.5 Pro to reproduce the first pages in several formats:

- Google Slides (Google Slides API)
- PowerPoint (python-pptx)
- HTML/CSS
- Quarto/Reveal.js

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