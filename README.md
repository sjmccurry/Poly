# Poly Document Generator

**Poly** is a lightweight Python-based document automation tool that replaces `{poly.Variable}` placeholders in Word `.docx` files with user input — while preserving original formatting.

## Features

- Automatically detects variables like `{poly.ClientName}` across one or more templates
            '{poly.VariableVariable}'
- Clean modern GUI inspired by enterprise tools like Clio
- Supports uploading or selecting built-in templates
- Preserves all formatting (bold, underline, alignment, etc.)
- Generates filled documents with a single click

## Usage

1. Place `.docx` templates into the `templates` folder, or use the Upload button.
2. Launch the app by running `poly_gui.py`.
3. Select one or more templates.
4. Enter values for the detected variables in the sidebar.
5. Click **Generate Document(s)** to output filled copies.


## IMPORTANT
You MUST have a "templates" and a "output" folder in the same directory as the code files or else it will not work! 
ALWAYS run the poly_updater file!
