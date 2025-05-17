/* global document, Office, console */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("generateButton").onclick = handleClick;
    document.getElementById("insertButton").onclick = insertSlides;
  }
});

// for font size and font style
document.addEventListener("DOMContentLoaded", function () {
  const optionSelect = document.getElementById("optionSelect");

  optionSelect.addEventListener("change", function () {
    const selectedOption = optionSelect.value;

    // Hide all font options
    document.querySelectorAll(".font-options").forEach((element) => {
      element.style.display = "none";
    });

    // Show the selected option's font options
    if (selectedOption === "title") {
      document.getElementById("titleOptions").style.display = "block";
    } else if (selectedOption === "subtitle") {
      document.getElementById("subtitleOptions").style.display = "block";
    } else if (selectedOption === "text") {
      document.getElementById("textOptions").style.display = "block";
    }
  });
});

// To store generated slides as a list of dictionaries
let slideDictionaries = [];
let currentDictIndex = 0;

async function handleClick() {
  document.getElementById("statusMessage").innerHTML = "Please wait, generating slides...";

  let userInput = document.getElementById("textInput").value;
  let slideInput = document.getElementById("slideInput").value;
  console.log("User input: ", userInput); // Debug
  userInput =
    "You are an expert in making professional MS PowerPoint Presentation on this topic: " +
    userInput +
    "Create" +
    slideInput +
    "slides, Frist slide contain Topic name and presentator name and Last slide contain message:Thank you for your time.I hope this presentation has provided you with a comprehensive overview. All slides contain Title, subtitle, content and images except frist and last. Note: Return only valid slides by slides. Please do not add Speaker Notes";
  console.log("Modified user input for AI: ", userInput); // Debug

  // Call the function to delete the first slide and add a blank slide
  await deleteFirstSlideAndAddBlank();
  // Generate all slides
  const slides = await generateSlides(userInput);
  console.log("Generated slides: ", slides); // Debug

  // Split slides into dictionaries
  slideDictionaries = splitSlidesIntoDictionaries(slides);
  console.log("Slides split into dictionaries: ", slideDictionaries); // Debug

  // Show the insert button only after all slides are generated and stored
  document.getElementById("insertButton").style.display = "block";
  document.getElementById("statusMessage").innerHTML = "Slides generated. Click 'Insert' to start inserting slides.";
}

// Function to split slides into separate dictionaries
function splitSlidesIntoDictionaries(slides) {
  const dictionaries = [];
  const listSize = 1; // Number of slides per dictionary; adjust as needed

  for (let i = 0; i < slides.length; i += listSize) {
    let slide = slides.slice(i, i + listSize)[0];
    console.log(`Processing slide ${i + 1}: `, slide); // Debug

    // Clean up trailing or misplaced asterisks before pushing the slide
    slide = cleanAsteriskIssues(slide);
    console.log(`Cleaned slide ${i + 1}: `, slide); // Debug

    dictionaries.push({ slideNumber: i + 1, content: slide });
  }

  console.log("Final dictionaries: ", dictionaries); // Debug
  return dictionaries;
}

// Function to clean up asterisk-related formatting issues
function cleanAsteriskIssues(slide) {
  console.log("Cleaning slide of asterisk issues: ", slide); // Debug

  // Look for any instance where ** is found only at the end of the slide and not at the start
  slide = slide.replace(/\*\*$/, ""); // Remove trailing ** from the end of the slide

  // Ensure all bold markers are properly balanced (remove any that are not)
  slide = slide.replace(/\*\*(?!\w)|(?<!\w)\*\*/g, ""); // Remove any lone asterisks

  console.log("Cleaned asterisks slide: ", slide); // Debug
  return slide;
}

// Function to start inserting slides one by one
function insertSlides() {
  console.log("Insert button clicked."); // Debug
  console.log("Current dictionary index: ", currentDictIndex); // Debug

  if (slideDictionaries.length > 0 && currentDictIndex < slideDictionaries.length) {
    console.log("Inserting slides from dictionary index: ", currentDictIndex); // Debug
    insertSlidesFromDictionary(slideDictionaries[currentDictIndex], function () {
      if (currentDictIndex + 1 < slideDictionaries.length) {
        currentDictIndex++;
        insertSlides(); // Move to the next dictionary
      } else {
        document.getElementById("statusMessage").innerHTML = "All slides have been inserted!";
      }
    });
  } else {
    document.getElementById("statusMessage").innerHTML = "No slides available to insert!";
    console.log("No slides to insert."); // Debug
  }
}

//Function to insert slides from a single dictionary
async function insertSlidesFromDictionary(slideDict, callback) {
  console.log("Inserting slide from dictionary: ", slideDict); // Debug

  if (slideDict && slideDict.content) {
    await insertTextIntoSlides(slideDict.content, async function () {
      await addSlideWithDelay(); // Ensure the slide is fully added with a delay before moving on
      callback(); // Finish inserting the slide
    });
  } else {
    console.log("No content available in the dictionary."); // Debug
    callback(); // Finish inserting if no content
  }
}

//Updated function to include user-defined font size and font style
async function addTextBox(slide, text, position, fontSize, isBold = false, fontName = "Arial") {
  await PowerPoint.run(async (context) => {
    // Add a text box to the slide with the specified position and size
    const textBox = slide.shapes.addTextBox(text, {
      left: position.left,
      top: position.top,
      width: position.width,
      height: position.height,
    });

    // Ensure the text box is fully loaded before applying formatting
    await context.sync();

    // Apply font size and line spacing
    const textFrame = textBox.textFrame;
    const textRange = textFrame.textRange;

    // Set font size for all text
    textRange.font.size = fontSize;

    // Apply line spacing if needed
    textFrame.textRange.paragraphFormat.lineSpacing = fontSize * 1.2;

    // Apply bold formatting if specified
    if (isBold) {
      textRange.font.bold = true;
    }

    // Apply font name
    textRange.font.name = fontName;

    await context.sync();

    console.log(
      `Text box added with font size ${fontSize}, line spacing ${fontSize * 1.2}, bold=${isBold}, fontName=${fontName}.`
    );
  });
}
async function addCenteredBoldTextBox(slide, text, position, fontSize, isBold = false, fontName = "Arial") {
  await PowerPoint.run(async (context) => {
    // Add a text box to the slide with the specified position and size
    const textBox = slide.shapes.addTextBox(text, {
      left: position.left,
      top: position.top,
      width: position.width,
      height: position.height,
    });

    // Ensure the text box is fully loaded before applying formatting
    await context.sync();

    // Apply font size and line spacing
    const textFrame = textBox.textFrame;
    const textRange = textFrame.textRange;

    // Set font size for all text
    textRange.font.size = fontSize;

    // Apply line spacing if needed
    textFrame.textRange.paragraphFormat.lineSpacing = fontSize * 1.2;

    // Apply bold formatting if specified
    if (isBold) {
      textRange.font.bold = true;
    }

    // Apply font name
    textRange.font.name = fontName;
    textFrame.textRange.paragraphFormat.horizontalAlignment = "Center";
    await context.sync();

    console.log(
      `Text box added with font size ${fontSize}, line spacing ${fontSize * 1.2}, bold=${isBold}, fontName=${fontName}.`
    );
  });
}

async function insertTextIntoSlides(text, callback) {
  console.log("Inserting text into slide: ", text); // Debug

  try {
    // Get user-defined font size and style
    const fontSizeTitleElement = document.getElementById("titleFontSizeSelect");
    const fontNameTitleElement = document.getElementById("titleFontStyleSelect");

    const fontSizeSubtitleElement = document.getElementById("subtitleFontSizeSelect");
    const fontNameSubtitleElement = document.getElementById("subtitleFontStyleSelect");

    const fontSizeOtherElement = document.getElementById("otherTextFontSizeSelect");
    const fontNameOtherElement = document.getElementById("otherTextFontStyleSelect");

    if (
      !fontSizeTitleElement ||
      !fontNameTitleElement ||
      !fontSizeSubtitleElement ||
      !fontNameSubtitleElement ||
      !fontSizeOtherElement ||
      !fontNameOtherElement
    ) {
      console.error("One or more font size/style elements are missing.");
      return;
    }

    const fontSizeTitle = parseInt(fontSizeTitleElement.value, 10) || 28;
    const fontSizeSubtitle = parseInt(fontSizeSubtitleElement.value, 10) || 24;
    const fontSizeOther = parseInt(fontSizeOtherElement.value, 10) || 20;

    const fontNameTitle = fontNameTitleElement.value || "Arial";
    const fontNameSubtitle = fontNameSubtitleElement.value || "Arial";
    const fontNameOther = fontNameOtherElement.value || "Arial";

    await PowerPoint.run(async (context) => {
      // Get the selected slides
      const slides = context.presentation.getSelectedSlides();
      slides.load("items");
      await context.sync();

      if (slides.items.length > 0) {
        let slide = slides.items[0]; // Use the first selected slide

        // Split the text into lines
        const lines = text.split("\n");

        let titleText = "";
        let subtitleText = "";
        let otherText = "";

        // Define the position of each text box
        const titlePosition = { left: 40, top: 40, width: 880, height: 50 };
        const subtitlePosition = { left: 40, top: 100, width: 880, height: 50 };
        const otherTextPosition = { left: 40, top: 140, width: 880, height: 360 };
        const titlePosition1 = { left: 40, top: 180, width: 880, height: 50 };
        const otherTextPosition1 = { left: 40, top: 250, width: 880, height: 100 };

        //             // Keywords to remove from text
        // Updated keywords list to include bullet point prefixes
        const keywordsToRemove = [
          "Title:", "Heading:", "Subtitle:", "Subheading:", 
          "• Title:", "• Heading:", "• Subtitle:", "• Subheading:", 
          "• Content:", "• Image:"
        ];

        // Separate text into title, subtitle, and other categories and remove keywords
        lines.forEach((line) => {
          const trimmedLine = line.trim();
          
          // Handle bullet-point title formats
          if (trimmedLine.startsWith("• Title:") || trimmedLine.startsWith("•Title:")) {
            titleText += trimmedLine.replace(/^• Title:|^•Title:/, "").trim() + "\n";
          } 
          // Handle bullet-point subtitle formats
          else if (trimmedLine.startsWith("• Subtitle:") || trimmedLine.startsWith("•Subtitle:")) {
            subtitleText += trimmedLine.replace(/^• Subtitle:|^•Subtitle:/, "").trim() + "\n";
          }
          // Handle bullet-point content formats - strip prefix completely 
          else if (trimmedLine.startsWith("• Content:") || trimmedLine.startsWith("•Content:")) {
            otherText += trimmedLine.replace(/^• Content:|^•Content:/, "").trim() + "\n";
          }
          // Handle bullet-point image formats - strip prefix completely
          else if (trimmedLine.startsWith("• Image:") || trimmedLine.startsWith("•Image:")) {
            otherText += trimmedLine.replace(/^• Image:|^•Image:/, "").trim() + "\n";
          } else {
            otherText += line.trim() + "\n";
          }
        });

        // Add the centered bold text box if there is only one slide
        if (currentDictIndex == 1 || currentDictIndex == slideDictionaries.length - 1) {
          if (titleText) {
            console.log("Adding title text box.");
            await addCenteredBoldTextBox(slide, titleText.trim(), titlePosition1, fontSizeTitle, true, fontNameTitle);
          }
          // Add the subtitle text box if there is content
          if (otherText) {
            console.log("Adding subtitle text box.");
            await addCenteredBoldTextBox(slide, otherText.trim(), otherTextPosition1, fontSizeOther, fontNameOther);
          }
        } else {
          // Add the title text box if there is content
          if (titleText) {
            console.log("Adding title text box.");
            await addTextBox(slide, titleText.trim(), titlePosition, fontSizeTitle, true, fontNameTitle);
          }

          // Add the subtitle text box if there is content
          if (subtitleText) {
            console.log("Adding subtitle text box.");
            await addTextBox(slide, subtitleText.trim(), subtitlePosition, fontSizeSubtitle, true, fontNameSubtitle);
          }

          // Add the other text box if there is content and handle splitting
          // Handle the other text
          const otherTextLines = otherText.trim().split("\n");
          if (otherTextLines.length > 10) {
            // Split into chunks of 10 lines
            const firstChunk = otherTextLines.slice(0, 10).join("\n");
            const remainingChunk = otherTextLines.slice(10).join("\n");
            // Add the first chunk of text to the current slide
            console.log("Adding first chunk of other text box.");
            await addTextBox(slide, firstChunk.trim(), otherTextPosition, fontSizeOther, false, fontNameOther);
            console.log("Done first chunk of other text box.");
            // Ensure that the slide is fully updated before adding a new one
            await context.sync();

            // Add a delay of 500ms before adding a new slide
            await new Promise((resolve) => setTimeout(resolve, 500));
            // Add a new slide and then insert the remaining chunk of text
            await addSlideWithDelay();

            // After the new slide is added, get the new slide and add the remaining text
            await PowerPoint.run(async (context) => {
              const slides = context.presentation.slides;
              slides.load("items");
              await context.sync();

              const newSlide = slides.items[slides.items.length - 1]; // Get the last added slide

              console.log("Adding remaining chunk of other text box to new slide.");
              await addTextBox(newSlide, remainingChunk.trim(), otherTextPosition, fontSizeOther, false, fontNameOther);
            });
          } else {
            // Add the other text box if there are 10 or fewer lines
            console.log("Adding other text box with 10 or fewer lines.");
            await addTextBox(slide, otherText.trim(), otherTextPosition, fontSizeOther, false, fontNameOther);
          }
        }
        callback();
      }
    });
  } catch (error) {
    console.error("Error inserting text into slides:", error);
  }
}

// Function to add a blank slide to PowerPoint and move to it with a delay
async function addSlideWithDelay() {
  console.log("Adding blank slide."); // Debug

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    const blankLayout = slides.getItemAt(0).layout; // This will usually point to the first layout

    await context.sync(); // Wait for the slide information to be retrieved

    // Add a new slide using the blank layout if available
    slides.add({
      layoutId: blankLayout.id, // This retrieves the first layout, often blank
    });

    await context.sync(); // Wait for the slide to be fully added before moving to the next step
    await goToLastSlide(); // Wait for the navigation to the last slide

    // Add a delay to ensure that the slide navigation is fully processed
    await new Promise((resolve) => setTimeout(resolve, 500)); // 500ms delay
  });
}

// Function to navigate to the last slide after adding it
async function goToLastSlide() {
  console.log("Navigating to last slide."); // Debug

  return new Promise((resolve, reject) => {
    Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Index, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Error navigating to last slide: ", asyncResult.error.message); // Debug
        setMessage("Error: " + asyncResult.error.message);
        reject(asyncResult.error);
      } else {
        resolve(); // Resolve the promise when navigation is complete
      }
    });
  });
}
//deleteFirstSlideAndAddBlank
// Function to delete the first slide and add a blank slide
// Function to delete the first slide and add a blank slide
// Function to delete the first slide and add a blank slide
// Function to delete the first slide and add a blank slide
// Function to delete the first slide and add a blank slide
async function deleteFirstSlideAndAddBlank() {
  console.log("Deleting the first slide and adding a blank slide."); // Debug

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    if (slides.items.length > 0) {
      // Delete the first slide
      slides.items[0].delete();
      await context.sync();
      console.log("First slide deleted.");

      // Add a new slide using the default layout
      slides.add(); // Default layout, no shape manipulation
      await context.sync();
      console.log("New default slide added.");
    } else {
      console.error("No slides available to delete.");
    }
  }).catch((error) => {
    console.error("Error in deleteFirstSlideAndAddBlank function:", error);
  });
}

function formatText(text) {
  console.log("Formatting text: ", text); // Debug

  // Remove "##" for headings
  text = text.replace(/##\s*/g, ""); // Removes "##" from headings

  // Convert **bold** to plain text
  text = text.replace(/\*\*(.*?)\*\*/g, "$1"); // Removes ** markers around bold text

  // Convert bullet points (handles both single-level and nested points)
  text = text.replace(/^\s*[-*•]\s+(.*)/gm, "\u2022 $1"); // Converts -, * or • to bullet points (•)

  // Handle nested bullet points (indented with two spaces for example)
  text = text.replace(/^\s{2}[-*•]\s+(.*)/gm, "\u2022    $1"); // Indents nested bullets

  // Remove specific keywords: "Text:", "Bullet Points:", "Content:", "Lists:", "List:"
  const keywordsToRemove = [
    "Text:", "Bullet Points:", "Content:", "Lists:", "List:", 
    "• Title:", "• Subtitle:", "• Content:", "• Image:", 
    "•Title:", "•Subtitle:", "•Content:", "•Image:"
  ];
  
  keywordsToRemove.forEach((keyword) => {
    text = text.replace(new RegExp(`^\\s*${keyword}\\s*`, "gm"), "");
  });

  // Remove any line containing "Slide X:" where X is a number, e.g., "Slide 8: Applications of Machine Learning"
  text = text.replace(/^\s*Slide\s*\d+\s*:.*(\r?\n)?/gm, "");

  console.log("Formatted text: ", text); // Debug
  return text.trim();
}

// Function to generate slides content using Google Generative AI
import { GoogleGenerativeAI } from "@google/generative-ai";

// Get API key from environment variable
const apiKey = process.env.GOOGLE_GEMINI_API_KEY || "";
if (!apiKey) {
  console.error("API key is missing! Make sure GOOGLE_GEMINI_API_KEY is set in .env file");
}

const genAI = new GoogleGenerativeAI(apiKey);
const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });

// Function to generate slides content using the AI model
async function generateSlides(userPrompt) {
  console.log("Generating slides with prompt: ", userPrompt); // Debug

  try {
    const result = await model.generateContent(userPrompt);
    const responseText = await result.response.text();

    console.log("AI response text: ", responseText); // Debug

    // Split the response based on "Slide" or some marker that separates slides
    const rawSlides = responseText.split(/(?=Slide\s\d+:)/); // Split by "Slide X: " but keep it as part of the split

    // Format each slide to remove extra line breaks and unnecessary markdown
    const formattedSlides = rawSlides.map((slide) => formatText(slide.trim()));

    console.log("Formatted slides: ", formattedSlides); // Debug
    return formattedSlides;
  } catch (error) {
    console.error("Error in generateSlides function: ", error); // Debug
    throw error;
  }
}

function setMessage(message) {
  console.log("Setting message: ", message); // Debug
  document.getElementById("statusMessage").innerText = message;
}
