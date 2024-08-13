## Python Setup Guide for Windows

## Installing Python
1. Download the latest Python installer from the official website `python.org`: Python Downloads.
2. Run the installer and follow the prompts.
3. Make sure to check the box that says "Add Python to PATH" during installation.

## Adding Python to PATH (Setting up Environment Variable)
1. Locate the directory where your Python executable (`python.exe`) lives. It could be in:
   - `C:\Python\`
   - Your `AppData\Local\Programs\Python` folder (e.g., `C:\Users\<USER>\AppData\Local\Programs\Python`)
   Replace `<USER>` with your logged-in username.
2. Verify that the executable works by double-clicking it and ensuring it opens a Python REPL.
3. To add Python to PATH:
   - Open System Properties (Win + Pause).
   - Go to the Advanced tab and click "Environment Variables."
   - Under System Variables, edit the `Path` variable.
   - Add the path to your Python directory (e.g., `C:\Python\`) at the end.
   - Restart Command Prompt for changes to take effect.


## Creating the Virtual Environment (Using pipenv)

### Installation

1 - Install Python and Pip:
- Make sure you have Python installed. You can check by running: `python --version`
If not installed, download the latest Python 3.x version from `python.org`

2 - Install Pipenv
- Use pip to install Pipenv: `pip install --user pipenv`. This user installation prevents system-wide package conflicts.

3 - Activate Pipenv:
- Navigate to your project directory (or create an empty one).
- Run:
    `pipenv shell`. This activates the virtual environment.
- OR: Navigate to the project folder in the terminal and run python scripts without activating virtual environment: e.g `pipenv run python ./excel_pdf.py`

4 (a) - Install Packages
- Inside the virtual environment, install packages using: `pipenv install <package_name>
- Replace <package_name> with the package name.

4 (b) - Install Packages (Using pipfile)
- Run: `pipenv install`

4 (c) - Install Packages (Using pipfile.lock)
- Run: `pipenv sync` 

5 - Manage Dependencies
- Pipenv creates a `Pipfile` and `Pipfile.lock` to manage dependencies.
- Use `pipenv install --dev` for development packages.
- Update packages with pipenv update.

### Usage

- Run your Python scripts within the activated virtual environment: `pipenv run python ./excel_pdf.py`


## Creating a Virtual Environment (Using venv)
1. Open Command Prompt.
2. Navigate to your project directory.
3. Run: `python -m venv myenv` (Replace `myenv` with your preferred environment name.)


4. Activate the virtual environment: `myenv\Scripts\activate`


## Installing Packages
1. With the virtual environment active, use `pip` to install packages: `pip install package-name`


## Deactivating the Virtual Environment
1. To exit the virtual environment, run: `deactivate`


That's it! You're all set up with Python on Windows. Happy coding! 🐍

### Note: Before running the code drop your excel file in the folder `input_xlsx`







