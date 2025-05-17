// This file contains instructions to fix the bullet point issue with text recognition

/* 
In the insertTextIntoSlides function, you need to modify the line parsing section to handle 
bullet-pointed sections with "• Title:", "• Subtitle:", etc.

Replace the code section:

```javascript
// Separate text into title, subtitle, and other categories and remove keywords
lines.forEach((line) => {
  if (line.startsWith("Title:") || line.startsWith("Heading:")) {
    titleText += line.replace(/^Title:|^Heading:/, "").trim() + "\n";
  } else if (line.startsWith("Subtitle:") || line.startsWith("Subheading:")) {
    subtitleText += line.replace(/^Subtitle:|^Subheading:/, "").trim() + "\n";
  } else {
    otherText += line.trim() + "\n";
  }
});
```

with:

```javascript
// Separate text into title, subtitle, and other categories and remove keywords
lines.forEach((line) => {
  const trimmedLine = line.trim();
  
  // Handle bullet-point title formats
  if (trimmedLine.startsWith("• Title:") || trimmedLine.startsWith("•Title:")) {
    titleText += trimmedLine.replace(/^• Title:|^•Title:/, "").trim() + "\n";
  } 
  // Handle standard title formats
  else if (trimmedLine.startsWith("Title:") || trimmedLine.startsWith("Heading:")) {
    titleText += trimmedLine.replace(/^Title:|^Heading:/, "").trim() + "\n";
  }
  // Handle bullet-point subtitle formats
  else if (trimmedLine.startsWith("• Subtitle:") || trimmedLine.startsWith("•Subtitle:")) {
    subtitleText += trimmedLine.replace(/^• Subtitle:|^•Subtitle:/, "").trim() + "\n";
  }
  // Handle standard subtitle formats
  else if (trimmedLine.startsWith("Subtitle:") || trimmedLine.startsWith("Subheading:")) {
    subtitleText += trimmedLine.replace(/^Subtitle:|^Subheading:/, "").trim() + "\n";
  }
  // Handle bullet-point content formats - strip prefix completely 
  else if (trimmedLine.startsWith("• Content:") || trimmedLine.startsWith("•Content:")) {
    otherText += trimmedLine.replace(/^• Content:|^•Content:/, "").trim() + "\n";
  }
  // Handle bullet-point image formats - strip prefix completely
  else if (trimmedLine.startsWith("• Image:") || trimmedLine.startsWith("•Image:")) {
    otherText += trimmedLine.replace(/^• Image:|^•Image:/, "").trim() + "\n";
  } else {
    otherText += trimmedLine + "\n";
  }
});
```
*/
