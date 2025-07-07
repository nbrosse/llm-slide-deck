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
