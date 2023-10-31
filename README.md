# Quiz Generator for Educational Use

**Author:** Mohammed Naser Alshehhi  
**Date:** 31st October 2023

## Introduction

The Quiz Generator is a software tool designed to assist educators in creating quizzes based on PowerPoint presentation content. It leverages the OpenAI GPT-3.5 model to generate a variety of question types and has been tested within the Higher Colleges of Technology.

## Installation

### Prerequisites

- Python 3.7 or higher
- pip (Python package installer)
- OpenAI API with a valid API key

### Dependencies

Install the required Python libraries using pip:

```bash
pip install python-pptx openai tk
```

### Installation Steps

1. Clone the repository or download the source code.
2. Navigate to the project directory.
3. Install the required dependencies.

## Usage

### Setting Up

Set your OpenAI API key in the script:

```python
openai.api_key = "Your_API_Key_Here"
```

### Running the Application

Execute the script:

```bash
python quiz_generator.py
```

### Generating Quizzes

1. Launch the application.
2. Load the PowerPoint file.
3. Enter the number of each question type.
4. Generate and copy the questions.

## Features

- Support for multiple question types.
- Custom question type support.
- Cognitive complexity consideration.
- User-friendly GUI.

## Testing

Extensive testing was conducted at the Higher Colleges of Technology.

## Contributing

To contribute, fork the repo, create a branch, commit your changes, and create a pull request.

## License

This project is licensed under the MIT License.

## Contact

For support, contact Mohammed Naser Alshehhi at mohammed.alshehi.95@gmail.com
