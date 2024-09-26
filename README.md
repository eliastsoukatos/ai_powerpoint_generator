# Interactive Presentation Generator with DALL-E 3 Images

This Python script creates PowerPoint presentations with AI-generated images using DALL-E 3. It allows users to input slide content interactively and generates relevant images for each slide based on the content and overall presentation theme.

## Features

- Interactive slide creation
- AI-generated images using DALL-E 3
- Automatic image prompt generation using GPT-4
- Customizable presentation saving options

## Prerequisites

- Python 3.7+
- Virtual environment (recommended)

## Installation

1. Clone this repository or download the script.

2. Create and activate a virtual environment:

```bash
python3 -m venv venv
source venv/bin/activate
```

3. Install the required packages:

```bash
pip install python-pptx python-dotenv openai requests
```

4. Create a `.env` file in the same directory as the script and add your OpenAI API key:

```
OPENAI_API_KEY=your_api_key_here
```

## Usage

1. Run the script:

```bash
python presentation_generator.py
```

2. Follow the prompts to add slides and create your presentation.

3. Enter a file name and location to save your presentation when prompted.

## Note

This script uses the OpenAI API, which may incur costs. Please review the [OpenAI pricing](https://openai.com/pricing) before using this script.

## License

This project is open-source and available under the MIT License.