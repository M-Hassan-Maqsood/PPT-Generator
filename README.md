# PowerPoint Slide Generator

An Office Add-in for PowerPoint that generates slides using Google Gemini AI based on user input.

## Features

- Generate slides based on user-provided topics
- Customizable font options for titles, subtitles, and body text
- AI-powered content generation using Google Gemini API
- Automatic slide formatting and layout

## Setup & Installation

### Prerequisites

- Node.js and npm
- Microsoft PowerPoint
- A Google Gemini API key

### Installation Steps

1. Clone the repository
2. Install dependencies:
   ```
   npm install
   ```
3. Set up environment variables:
   - Copy `.env.example` to `.env`
   - Add your Google Gemini API key to the `.env` file:
     ```
     GOOGLE_GEMINI_API_KEY=your_api_key_here
     ```

### Development

Run the add-in in development mode:

```
npm run start:desktop -- --app powerpoint
```

### Build

Build for production:

```
npm run build
```

## Usage

1. Open PowerPoint
2. Navigate to the add-in taskpane
3. Enter a topic and select the number of slides
4. Choose font options for titles, subtitles, and text
5. Click "Generate" to create slides
6. Click "Insert" to add the slides to your presentation

## License

This project is licensed under the MIT License.
