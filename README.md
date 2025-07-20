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



TODO Test quarto rendering powerpoint

### A. Marp (Markdown Presentation Ecosystem)

*   **How it Works:** You write your slides in Markdown, using `---` to separate them. You can add directives (e.g., `<!-- theme: gaia -->`) to control the appearance. The **Marp CLI** tool can then convert this `.md` file into a `.pptx` file.
*   **Example (`slides.md`):**
    ```markdown
    ---
    theme: uncover
    ---

    # My Awesome Report
    Generated from Markdown

    ---

    ## Slide 2: A Bulleted List
    - Point 1
    - Point 2
    - Point 3
    ```
*   **Command Line:**
    ```bash
    # Install: npm install -g @marp-team/marp-cli
    # Convert to PPTX:
    marp --pptx slides.md -o slides.pptx
    ```

### B. Pandoc

*   **How it Works:** Pandoc is a universal document converter. It can convert between dozens of formats. It's incredibly powerful for converting Markdown to `.pptx`, and you can even specify a reference `.pptx` file to use as a template for styling.
*   **Command Line:**
    ```bash
    # Simple conversion
    pandoc my_document.md -o my_presentation.pptx

    # Using a custom template for styling
    pandoc my_document.md --reference-doc=custom-template.pptx -o my_presentation.pptx
    ```