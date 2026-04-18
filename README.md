# FuncAtlas

FuncAtlas is a GUI-based analysis tool built with PySide6.

## Prerequisites

- Python 3.8 or higher installed on your system.

## Setup Instructions

Follow these steps to set up the project locally:

### 1. Create a Virtual Environment
Open your terminal in the project root directory and run:
```powershell
python -m venv .venv
```

### 2. Activate the Virtual Environment
Activate the environment to ensure dependencies are installed locally:

**Windows (PowerShell):**
```powershell
.\.venv\Scripts\Activate.ps1
```

**Windows (Command Prompt):**
```cmd
.\.venv\Scripts\activate
```


### 3. Install Dependencies
Install the required Python packages:
```powershell
pip install -r requirements.txt
```

## Running the Application

Once the virtual environment is activated and dependencies are installed, you can start the application by running:

```powershell
python main.py
```

## Project Structure
- `main.py`: The entry point of the application.
- `main_window.py`: Contains the main window and UI logic.
- `pages/`: Individual UI pages/screens.
- `ui/`: Custom widgets and UI components.
- `core/`: Core logic and backend functionality.
- `services/`: External services or helper logic.
