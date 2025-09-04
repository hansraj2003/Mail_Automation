# Mail_Automation
**Project Setup Guide**
This guide will walk you through the steps to set up and run the project.

Prerequisites
Before you begin, ensure you have the following installed on your system:

Git: For cloning the repository.

VS Code: The recommended code editor.

Python 3.x: The project is built with Python.

Getting Started
Follow these steps in order to set up the project locally.

Clone the Repository
Open your terminal or command prompt, navigate to the directory where you want to store the project, and clone the repository using the following command:

git clone <repository_url>

Replace <repository_url> with the actual URL of your Git repository.

Configure Environment Variables
This project requires environment variables to run.

Locate the envsample file in the root directory.

Create a copy of this file and rename it to .env.

Open the newly created .env file and fill in the required values. Do not change the original envsample file.

Personalize the Script

Open the updated_mailer.py file.

Use the search function (Ctrl + F or Cmd + F) to find all instances of the name "Hansraj".

Replace all occurrences with your name.

Set Up the Virtual Environment
It is recommended to use a virtual environment to manage project dependencies. Follow these steps:

Open the integrated terminal in VS Code.

Create the virtual environment:

python -m venv venv

Activate the virtual environment:

venv\Scripts\activate

Install the required packages from the requirements.txt file:

pip install -r requirements.txt

Run the Script
Once all the dependencies are installed, you can run the main script.

Ensure your virtual environment is activated.

Execute the script with the following command:

py updated_mailer.py

If the setup was successful, the command should run without any errors.