import requests
import os
from zipfile import ZipFile
import openpyxl

def download_and_unzip_github_repository(repo_url):
    # Extract username and repository name from the GitHub URL
    _, _, _, username, repository = repo_url.rstrip('/').split('/')

    # Create a zip file name based on the repository name
    zip_file_name = f"{username}_{repository}_master.zip"

    # Construct the GitHub API URL to get the zipball of the repository
    api_url = f"https://api.github.com/repos/{username}/{repository}/zipball/master"

    # Send a GET request to the GitHub API to download the zipball
    response = requests.get(api_url)

    if response.status_code == 200:
        # Save the zipball to a local file
        with open(zip_file_name, 'wb') as zip_file:
            zip_file.write(response.content)

        print(f"Repository downloaded successfully as {zip_file_name}")

        # Unzip the downloaded file
        with ZipFile(zip_file_name, 'r') as zip_ref:
            zip_ref.extractall()

        print(f"Repository unzipped successfully.")

        # Analyze the unzipped repository and create Excel sheet
        analyze_and_create_excel()
    else:
        print(f"Failed to download repository. Status code: {response.status_code}")

def analyze_and_create_excel():
    # Create a new Excel workbook and select the active sheet
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Set the header row
    sheet.append(["Python File Paths"])

    print("\nAnalyzing the unzipped repository:")
    for root, dirs, files in os.walk("."):
        # Exclude the script itself from analysis
        if os.path.abspath(root) != os.path.abspath("."):
            for file in files:
                # Check if the file has a .py extension
                if file.endswith(".py"):
                    file_path = os.path.join(root, file)
                    print(f"Python file found: {file_path}")
                    # Add the file path to the Excel sheet
                    sheet.append([file_path])

    # Save the Excel workbook
    excel_file_name = "python_files.xlsx"
    workbook.save(excel_file_name)

    print(f"\nExcel sheet created successfully: {excel_file_name}")

if __name__ == "__main__":
    # Get GitHub repository URL from the user
    github_url = input("Enter the GitHub repository URL: ")

    # Download, unzip, and analyze the repository
    download_and_unzip_github_repository(github_url)
