# Question Paper Generator

## Overview
The Question Paper Generator is a command-line tool designed to help educators and faculty members easily create structured question papers. The tool extracts questions from a question bank document, allows users to organize them into sections, and generates a formatted question paper based on predefined templates.

## Features
- Interactive command-line interface
- Import questions from existing DOCX question banks
- Extract question metadata (marks, CO, RBT)
- Organize questions into structured exam sections
- Total marks calculation and validation
- Formatted output using customizable templates
- Support for question subparts with proper labeling (a, b, c, etc.)

## Requirements
- Python 3.6+
- python-docx library

## Installation

### Step 1: Clone the repository
```bash
git clone https://github.com/prathamngundikere/question-paper-generator.git
cd question-paper-generator
```

### Step 2: Set up a virtual environment
```bash
# Create a virtual environment
python -m venv venv

# Activate the virtual environment
# On Windows
venv\Scripts\activate
# On macOS/Linux
source venv/bin/activate
```

### Step 3: Install dependencies
```bash
pip install -r requirements.txt
```

The `requirements.txt` file should include:
```
python-docx==0.8.11
```

## Usage

### Running the application
```bash
python question_paper_generator.py
```

### Workflow
1. **Prepare your files**:
   - Create a question bank DOCX file containing questions with marks, CO, and RBT information in the format: 
     `Question text (Marks: X) (CO: Y) (RBT: Z)`
   - Create a template DOCX file with a table containing 5 columns (Question No., Question, Marks, CO, RBT)

2. **Select question bank**:
   - The program will display available DOCX files in the current directory
   - Select your question bank file

3. **Select questions for each section**:
   - PART A - Question 1: Select questions totaling ~25 marks
   - PART A - Question 2: Select alternative questions totaling ~25 marks
   - PART B - Question 3: Select questions totaling ~25 marks
   - PART B - Question 4: Select alternative questions totaling ~25 marks

4. **Select template**:
   - Choose the template DOCX file for formatting the output

5. **Generate output**:
   - Provide a filename for the generated question paper
   - The program will create a formatted question paper with proper numbering and organization

## Question Paper Structure
The generated question paper follows this structure:
- **PART A**
  - Question 1 (with subparts 1a, 1b, etc.)
  - Question 2 (with subparts 2a, 2b, etc.)
- **PART B**
  - Question 3 (with subparts 3a, 3b, etc.)
  - Question 4 (with subparts 4a, 4b, etc.)

Students are expected to choose between Q1 and Q2, and between Q3 and Q4, with each section worth approximately 25 marks.

## Project Structure
```
question-paper-generator/
│
├── question_paper_generator.py     # Main application script
├── requirements.txt                # Python dependencies
├── README.md                       # This file
├── question_banks/                 # Example folder for question banks
│   └── sample_question_bank.docx
└── templates/                      # Example folder for templates
    └── question_paper_template.docx
```

More better don't keep it in the folder...

## Question Format
Questions in the question bank should follow this format:
```
Explain the concept of inheritance in OOP. (Marks: 5) (CO: 2) (RBT: 2)
```

Where:
- `(Marks: X)` indicates the marks allocated to the question
- `(CO: Y)` indicates the Course Outcome number
- `(RBT: Z)` indicates the Revised Bloom's Taxonomy level

## Template Requirements
The template file should contain a table with the following columns:
1. Question No.
2. Question
3. Marks
4. CO
5. RBT

The program will preserve the header row and populate the table with the selected questions.

## Contributing
Contributions are welcome! Please feel free to submit a Pull Request.

## License
This project is licensed under the Apache License - see the LICENSE file for details.