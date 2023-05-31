# darkSkyTimes

The darkSkyTimes repository allows users to easily determine the times when the sky is the darkest (i.e. between the times of astronomical twilight end and start, along with the absence of the moon for a desired location).
## Description

The purpose of darkSkyTimes is to help identify the times when there is an absence of astronomical twilight along with an absence of the moon. Based on user-defined location (lat/lon) and date range (YYYY/MM/DD), darkSkyTimes calculates and outputs an Excel file containing the moon set, astronomical twilight end, astronomical twilight start, and moon rise times for the entire date range inputted by the user. Ideally, darkSkyTimes can be used to determine the ideal times when the convergance of when the sun, astronomical twilight, and the moons presence in the sky are not visibl and indicate the most optimal times to view the night sky imaginable. 

## Getting Started
### Dependencies
* Python3
* pip
* pipenv

### Installing
These installation instructions are written for Terminal in Mac OS X

* Run Terminal by navigating to Applications -> Utilities -> Terminal. You can also run Terminal by first pressing Command + Space Bar on your Mac, then typing in 'Terminal' and then pressing enter. 

* Use Terminal to create a new folder on your local machine where the repository will be cloned. In this example, I will create a folder named 'darkSkyTimes' on my desktop by running these commands:
   ```bash
   cd Desktop
   mkdir darkSkyTimes

* Navigate to the project directory:
   ```bash
   cd darkSkyTimes
   
* Clone the repository to your local machine using the following command:
   ```bash
   git clone https://github.com/BobbyBoy101/darkSkyTimes.git
   
* If you are sam and it asks you for username enter:
   ```bash
   BobbyBoy101
   
* If you are sam and it asks you for password, I have set up a personal access token (classic) for youuuuu. Paste this as the password in order to git clone:
   ```bash
   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
   
* Create a virtual environment
   ```bash
   pipenv install

* Install the required dependencies by running the following command:
   ```bash
   pip install -r requirements.txt
   
### Usage

1. After successfully installing the dependencies, you can run the script to generate the Excel spreadsheet.
   ```bash
   python main.py
2. Follow the prompts to input the desired date range.
3. The generated Excel spreadsheet will be saved in the project directory.

### License
* My balls
   ```bash
   in your mouth

