# darkSkyTimes

The darkSkyTimes repository generates an Excel spreadsheet with precise astronomical data, enabling users to easily identify the darkest moments of the night sky by inputting a date range and latitude/longitude coordinates.

## Description
The darkSkyTimes repository empowers users to effortlessly determine the optimal moments of darkness during the night sky. By inputting a desired date range and latitude/longitude coordinates, this project generates an Excel spreadsheet that meticulously captures crucial astronomical information. Each entry in the spreadsheet corresponds to an individual day within the specified range and provides precise data on astronomical twilight start, astronomical twilight end, moon rise, and moon set. The primary objective of darkSkyTimes is to facilitate the identification of ideal periods when the absence of sunlight, astronomical twilight, and the moon's presence converge, resulting in the darkest skies imaginable.

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
   
* Create a virtual environment
   ```bash
   pipenv install

* Install the required dependencies by running the following command:
   ```bash
   pip install -r requirements.txt
