# llm-slide-deck

Leveraging LLM to enhance the creation of slide decks

This repository provides the source code featured in the blog post [LLM-Powered Slide Decks: A Comparison of Formats](https://nbrosse.github.io/posts/llm-slides/llm-slides.html), which explores how large language models can streamline and enhance the process of creating slide presentations in various formats.

## Requirements

To install the requirements, use [uv](https://docs.astral.sh/uv/): 

```bash
uv sync
```

## Google Slides

In order to use the Google Slides API, you need to set up a project in the Google Cloud Console and enable the Google Slides API. Follow these steps:

1. Go to the [Google Cloud Console](https://console.cloud.google.com/).
2. Create a new project or select an existing project.
3. Enable the Google Slides API and the Google Drive API for your project.
4. Create credentials (OAuth 2.0 client IDs) for your application.
5. Download the credentials JSON file and save it as `credentials.json` in the google directory `google/`.

Then, execute the following command to create the slide deck on Google Slides:

```bash
uv run python google/create_slide_deck.py
```