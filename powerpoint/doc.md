### Part 2: PowerPoint (via Office Scripts)

Office Scripts are the modern way to automate PowerPoint. They run in PowerPoint for the web and are written in TypeScript.

#### How to Access the API

1.  **Subscription:** You need a Microsoft 365 commercial or educational license that includes Office desktop apps (e.g., Business Standard, E3, E5).
2.  **Open PowerPoint for the Web:** Go to `powerpoint.office.com`.
3.  **Open the Automate Tab:** Create or open a presentation, then click on the "Automate" tab in the ribbon.
4.  **Code Editor:** Click "New Script". This opens a code editor panel directly in PowerPoint. You can write, run, and save your scripts here.

#### Example: Creating the First 3 Slides in TypeScript

Paste this code into the Office Scripts editor in PowerPoint for the web and click "Run".

**Important Note on Images:** Office Scripts cannot access external URLs for security reasons. You must convert your image to a **Base64 encoded string** and paste it into the script. You can use an online converter for this. I've used a placeholder below.

```typescript
async function main(presentation: PowerPointScript.Presentation) {

  // --- SLIDE 1: TITLE SLIDE ---
  // A Base64 string of your background image. This is a tiny red dot as a placeholder.
  // To get your real image, use an online "image to base64" converter.
  const base64Image = "iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==";

  // Get the first slide
  let slide1 = presentation.getSlides()[0];
  
  // Clear any existing shapes on the first slide
  let shapes = slide1.getShapes();
  for (let shape of shapes) {
    shape.delete();
  }
  
  // Set the background image
  slide1.setBackground(base64Image, PowerPointScript.PictureFormat.PNG);

  // Add a text box for the title
  let titleBox = slide1.addTextBox("EDF PRESENTION\nAPRIL 2024", { x: 50, y: 200, width: 400, height: 100 });
  
  // Style the text
  let titleText = titleBox.getTextFrame().getTextRange();
  titleText.getFont().setBold(true);
  titleText.getFont().setSize(32);
  titleText.getFont().setColor("#FFFFFF"); // White color
  
  // --- SLIDE 2: DISCLAIMER SLIDE ---
  // Add a new slide with a 'Title and Content' layout
  let slide2 = presentation.addSlide(PowerPointScript.SlideLayout.TitleAndContent);

  // Get the title and body shapes from the layout
  let titleShape2 = slide2.getShape(0); // Assumes title is the first shape
  let bodyShape2 = slide2.getShape(1);  // Assumes body is the second shape

  // Set the text
  titleShape2.getTextFrame().setText("DISCLAIMER");
  bodyShape2.getTextFrame().setText("This presentation is for information purposes only and does not constitute or form part of a prospectus...");
  
  // Adjust font size for the body
  bodyShape2.getTextFrame().getTextRange().getFont().setSize(8);

  // --- SLIDE 3: INTRO SLIDE ---
  // Add a blank slide
  let slide3 = presentation.addSlide(PowerPointScript.SlideLayout.Blank);
  
  // Set the background image
  slide3.setBackground(base64Image, PowerPointScript.PictureFormat.PNG);
  
  // Add the white content box with a blue border
  // Note: Can't create the diagonal shape easily, so we'll make a styled rectangle
  let contentBox = slide3.addGeometricShape(PowerPointScript.GeometricShapeType.Rectangle, {x: 20, y: 80, width: 400, height: 400 });
  
  // Style the box
  contentBox.getFill().setSolidColor("#FFFFFF"); // White fill
  let border = contentBox.getLineFormat();
  border.setColor("#0078D4"); // A nice blue
  border.setWeight(8);
  
  // Add and style text inside the box
  contentBox.getTextFrame().setText("EDF PRESENTATION\n\nA world leader in generating carbon-free electricity, constantly available on demand");
  let contentText = contentBox.getTextFrame().getTextRange();
  contentText.getFont().setSize(18);
  contentText.getFont().setColor("#000000"); // Black text
}
```




# Building Slide Decks with Microsoft PowerPoint {#sec-microsoft-powerpoint}

Of course. Interacting with Microsoft PowerPoint through code is a very common requirement for business automation, reporting, and content generation. The possibilities are vast, but they fall into three distinct categories, each with its own ecosystem, strengths, and weaknesses.

Let's use an analogy: Imagine PowerPoint is a car.

1.  **COM Automation:** This is like sitting in the driver's seat and having a robot (your code) physically press the pedals, turn the wheel, and flip the switches. You have access to *everything* a human driver does.
2.  **Open XML SDK (File Manipulation):** This is like being a mechanic in a garage with the car turned off. You can open the hood, disassemble the engine, change the parts (the `.pptx` file contents), and put it all back together. You can't *drive* the car (render a PDF/video), but you can fundamentally change its structure.
3.  **Microsoft Graph API:** This is like using a modern smartphone app to remotely interact with your smart car. You can lock/unlock the doors, check the fuel level, and maybe get its location (file-level operations), but you can't perform fine-grained driving maneuvers.

### Method 1: COM Automation (The "Local Control" Method)

This is the traditional and most powerful method for controlling the PowerPoint application directly. It requires the PowerPoint desktop application to be installed on the machine where the code is running.

**What it is:** COM (Component Object Model) is a Microsoft technology that allows applications to expose their functionality to be controlled by other programs. Your code gets a "handle" to the running PowerPoint application and sends it commands.

**How it Works:**
*   **From within PowerPoint (VBA):** You can write Visual Basic for Applications (VBA) macros directly inside PowerPoint. This is the simplest way to start. You open the VBA editor (`Alt + F11`) and write code that manipulates the `ActivePresentation`.
*   **From an external application:** Languages like Python (using the `pywin32` library), C#, PowerShell, or VBScript can instantiate the PowerPoint application object, making it run (visibly or invisibly in the background) and then send it commands.

**Key Capabilities (Virtually Unlimited):**
*   **Full Presentation & Slide Control:** Create, open, save, and close presentations. Add, delete, duplicate, and reorder slides.
*   **Deep Shape & Content Manipulation:** Add any shape (text boxes, pictures, tables, charts, videos). Precisely control their size, position, rotation, and formatting (fill, line, shadow, etc.).
*   **Text & Data Integration:** Insert and format text, create bulleted lists, and populate tables and charts with data from other sources (like Excel).
*   **Run Animations & Transitions:** Programmatically apply and modify slide transitions and object animations.
*   **Exporting & Rendering:** **This is a key advantage.** Because you are controlling the actual application, you can command it to export the presentation to other formats like **PDF, JPG, PNG, or even MP4 video**.
*   **Run Macros & Interact with Add-ins:** You can trigger other macros or interact with installed third-party add-ins.

**Pros:**
*   **Unmatched Power & Flexibility:** Can do virtually anything a human user can do through the UI.
*   **Access to Rendering Engine:** The only reliable method to programmatically create PDFs or images from slides.

**Cons:**
*   **Windows Only:** COM is a Windows-specific technology. This will not work on Linux or macOS.
*   **Requires PowerPoint Installation:** The machine running the code must have a licensed copy of PowerPoint installed.
*   **Not Suitable for Servers:** Automating Office applications on a server (like an ASP.NET web server) is officially discouraged by Microsoft. It can be unstable, slow, and prone to hanging or showing dialog boxes that require user intervention.

**Example (Python using `pywin32`):**

```python
import win32com.client
import os

# Create an instance of the PowerPoint application
powerpoint = win32com.client.Dispatch("Powerpoint.Application")
powerpoint.Visible = 1 # Make the application visible

# Create a new presentation
presentation = powerpoint.Presentations.Add()

# Add a title slide (layout index 1 is typically title slide)
slide1 = presentation.Slides.Add(1, 1) 

# Set the title and subtitle
title_shape = slide1.Shapes.Title
subtitle_shape = slide1.Shapes(2) # Access the second shape (usually the subtitle)
title_shape.TextFrame.TextRange.Text = "Automated Report"
subtitle_shape.TextFrame.TextRange.Text = "Generated by Python"

# Save as a PPTX
ppt_path = os.path.join(os.getcwd(), "automated_report.pptx")
presentation.SaveAs(ppt_path)

# Save as a PDF (ppFormatPDF = 32)
pdf_path = os.path.join(os.getcwd(), "automated_report.pdf")
presentation.SaveAs(pdf_path, 32)

# Clean up
presentation.Close()
powerpoint.Quit()
```

### Method 2: Open XML SDK (The "Server-Side" Method)

This method involves directly manipulating the PowerPoint file (`.pptx`) itself, without ever launching the PowerPoint application.

**What it is:** A modern Office file (like `.pptx`, `.docx`, `.xlsx`) is not a single binary file. It is a **ZIP archive** containing a collection of XML files and media assets (images, etc.). The Open XML SDK provides libraries to programmatically read, write, and modify these internal parts.

**How it Works:**
Your code uses a library (like `python-pptx` for Python or the official Open XML SDK for .NET) to "unzip" the `.pptx` in memory. The library gives you a high-level API to interact with the XML structure (e.g., `presentation.slides.add_slide()`). When you save, the library re-packages all the modified XML and media files back into a valid `.pptx` ZIP archive.

**Key Capabilities:**
*   **Platform Independent Generation:** Create full presentations from scratch on any operating system (Windows, Linux, macOS).
*   **Template-Based Generation:** A very common use case is to start with a template `.pptx` file (with your company's branding, master slides, and layouts) and then programmatically add new slides and populate them with content.
*   **Content Manipulation:** Insert text, images, tables, and charts. Find and replace text placeholders.
*   **Read & Extract Data:** Parse an existing presentation to extract all text content, speaker notes, or images.

**Pros:**
*   **Cross-Platform:** Works on Windows, Linux, and macOS.
*   **No PowerPoint Required:** Does not need PowerPoint to be installed.
*   **Server-Friendly & Scalable:** Ideal for web applications and backend services that generate reports. It's fast and stable.

**Cons:**
*   **No Rendering Engine:** It cannot convert a slide to a PDF or an image, because it doesn't know how to *render* the content. It only manipulates the file structure.
*   **Limited Functionality:** It cannot run macros, apply certain complex effects, or interact with features that are not explicitly defined in the Open XML standard. Chart manipulation can be complex.
*   **Steeper Learning Curve:** You are working at a lower level and sometimes need to understand the underlying XML schema.

**Example (Python using `python-pptx`):**

```python
# pip install python-pptx
from pptx import Presentation
from pptx.util import Inches

# Create a new presentation (or open an existing one)
prs = Presentation() 

# Use a built-in layout (layout 5 is title and content)
title_and_content_layout = prs.slide_layouts[5]
slide = prs.slides.add_slide(title_and_content_layout)

# Set the title and content
title = slide.shapes.title
body = slide.placeholders[1] # Access the body placeholder
title.text = "Server-Generated Slide"
body.text = "This was created on a server without PowerPoint!"

# Add an image
img_path = 'logo.png'
slide.shapes.add_picture(img_path, Inches(1), Inches(3), height=Inches(1))

# Save the presentation
prs.save("server_generated_report.pptx")
```

### Method 3: Microsoft Graph API (The "Cloud" Method)

This is the modern, web-first approach for interacting with files stored in Microsoft 365 (OneDrive for Business or SharePoint).

**What it is:** A RESTful web API that provides a single endpoint (`https://graph.microsoft.com`) to access data across the Microsoft 365 ecosystem.

**How it Works:**
Your application (which can be anywhere) authenticates with Azure Active Directory using OAuth 2.0 to get an access token. It then makes standard HTTP requests (GET, POST, PATCH, DELETE) to the Graph API, specifying the user, the drive, and the file it wants to interact with.

**Key Capabilities (Currently Limited for PowerPoint Content):**
*   **File Management:** The primary use case. Upload, download, copy, move, and delete `.pptx` files in OneDrive/SharePoint.
*   **Permissions:** Manage sharing and permissions on a presentation file.
*   **Get Basic Information:** Retrieve metadata about the presentation.
*   **Limited Content Interaction:** The Graph API for PowerPoint is far less developed than the Google Slides API. As of late 2023, it does **not** offer deep, granular control to add/modify individual shapes, text, or tables within a slide.
*   **Cloud Rendering:** A key advantage is that you can **request a cloud-based conversion** of the file to formats like **PDF or thumbnails** without having PowerPoint installed locally.

**Pros:**
*   **Cloud-Native & Language Agnostic:** Can be called from any application that can make HTTP requests.
*   **Secure:** Uses modern, standard OAuth 2.0 authentication.
*   **Integrated with Microsoft 365:** The best way to automate workflows for files already living in SharePoint or OneDrive.

**Cons:**
*   **Very Limited Content Manipulation:** You cannot use it to build a presentation from scratch or perform detailed edits. Its capabilities are mostly at the file level.
*   **Requires Files in the Cloud:** Only works for presentations stored in OneDrive or SharePoint.
*   **Complex Authentication:** Setting up Azure AD app registration and the OAuth 2.0 flow can be complex.

**Example (Conceptual HTTP Request):**

```http
# Request to get a presentation's content as a PDF
# Note: You would first need to get a valid {access-token}

GET https://graph.microsoft.com/v1.0/me/drive/items/{item-id}/content?format=pdf
Authorization: Bearer {access-token}
```

### Summary Comparison Table

| Feature | COM Automation (e.g., pywin32) | Open XML SDK (e.g., python-pptx) | Microsoft Graph API |
| :--- | :--- | :--- | :--- |
| **Environment** | Windows Desktop | Any Server/Desktop (Windows, Linux, Mac) | Any (Cloud-based REST API) |
| **PowerPoint Required?** | **Yes** | **No** | **No** |
| **Primary Use Case**| Full automation of the app, reporting with PDF/Image export | Server-side generation of `.pptx` files, templating | Cloud file management, simple conversions |
| **Content Creation** | **Excellent.** Full control. | **Very Good.** Deep control over file content. | **Poor.** Very limited content APIs. |
| **Export to PDF/Image** | **Yes** (via the app's engine) | **No** (no rendering engine) | **Yes** (via cloud service) |
| **Robustness** | Medium (can hang on UI dialogs) | **High** (stable for server use) | **High** (stable API) |
| **Setup Complexity** | Low (if Python/VBA is installed) | Low (install a library) | **High** (Azure AD app registration) |

### Which Method Should You Choose?

*   **Choose COM Automation if:**
    *   You are working on a **Windows desktop**.
    *   You **absolutely need to export to PDF, images, or video**.
    *   You need to manipulate charts extensively or interact with third-party add-ins.

*   **Choose Open XML SDK if:**
    *   You are building a **web application or backend service**.
    *   Your code needs to run on **Linux or macOS**.
    *   Scalability and stability are critical, and you only need to output a `.pptx` file.

*   **Choose Microsoft Graph API if:**
    *   Your entire workflow is **based in Microsoft 365 (SharePoint/OneDrive)**.
    *   Your primary need is to manage files, permissions, or trigger a cloud-based conversion of an *existing* file.