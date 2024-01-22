# darkSkyTimes

The darkSkyTimes repository allows users to easily determine the times when the sky is the darkest (i.e. between the times of astronomical twilight end and start, along with the absence of the moon for a desired location).

## Description

The darkSkyTimes repository empowers users to effortlessly determine the optimal moments of darkness during the night sky. By inputting a desired date range and latitude/longitude coordinates, this project generates an Excel spreadsheet that meticulously captures crucial astronomical information. Each entry in the spreadsheet corresponds to an individual day within the specified range and provides precise data on astronomical twilight start, astronomical twilight end, moon rise, and moon set. The primary objective of darkSkyTimes is to facilitate the identification of ideal periods when the absence of sunlight, astronomical twilight, and the moon's presence converge, resulting in the darkest skies imaginable.

***Where this project is now***

Currently this project generates a .xlsx file that contains two sheets. The first sheet contains the ideal dark sky times and duration for the date range and location specified. The second sheet contains all of the raw data used to calculate this information.

Calculated dark sky times example:

![Current darkSkyTimes](CalculatedOutput.jpg)

Raw data output example:

![Current darkSkyTimes](RawData.jpg)

***KNOWN BUG: Generating data in a different time zone***

Please note that the data will be generated in the time zone where your computer is located. If the latitude/longitude you enter is in your local time zone, then the data will be displayed correctly. However, if you enter a latitude/longitude that is in a different time zone from where you generate the data, that data will be displayed in your local time zone and you must offset the data accordingly. For example, if you are in MST and want to calculate data from a location in PST, the data displayed will be an hour ahead of the actual time for that location. I am working on solving this problem and will update the code soon to allow the user to display the data in the correct time zone based on the location entered.

***DISCLAIMER: Work in Progress***

*Please be aware that this project is a work in progress and is not considered complete at this stage. It may be undergoing active development, and significant changes could occur. The code, documentation, and other project assets presented here are subject to continuous improvements, updates, and enhancements. As such, I cannot guarantee the stability, reliability, or completeness of the current state of the project. I strongly advise against using this project in a production environment or for critical tasks at this time. It may contain bugs, security vulnerabilities, or incomplete features that could lead to unexpected behavior or data loss.*

## Getting Started
### Dependencies
* Python3
* pip
* pipenv

### Installing
These installation instructions are written for Terminal in Mac OS X

* Run Terminal by navigating to Applications -> Utilities -> Terminal. You can also run Terminal by first pressing Command + Space Bar on your Mac, then typing in 'Terminal' and then pressing enter. 

* Use Terminal to create a new folder on your local machine where the repository will be cloned. In this example, I will create a folder named 'Dark_Sky_Times' on my desktop by running these commands:
   ```bash
   cd Desktop
   mkdir Dark_Sky_Times

* Navigate to the new folder that you just created on your desktop project directory:
   ```bash
   cd Dark_Sky_Times
   
* Clone the repository to your local machine using the following command:
   ```bash
   git clone https://github.com/BobbyBoy101/darkSkyTimes.git

* Navigate to the project directory:
   ```bash
   cd darkSkyTimes
   
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
2. Follow the prompts to input the desired month and latitude/longitude coordinates.
3. The generated Excel spreadsheet will be saved in the project directory.

### Parameters

Here is an example parameter:

Please note that the -tz parameter is currently non-functioning, but will be fixed soon.

~~~
-lat 40.7720 -lon -112.1012 -start 04/01/2023 -end 12/31/2023 -tz America/Denver
~~~

### License
This project is licensed under the [MIT License](LICENSE)

