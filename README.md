# FSI-Python by Jenny Liu

## BACKGROUND ##
This set of scripts is intended to be used alongside the Floc Strength Instrument created by Instrument Works for image processing to determine the size of flocs in water from different sources. It aims to first identify particles, then take various approaches to measure floc size based on area, diameter, sphericity, and more. The intent is to offer the ability to test a multitude of images and flocs at once and have the results exported to an excel spreadsheet. If the flowcell becomes dirty or repeated particles show up, another script can be run on the resulting excel file to eliminate such repeats so that results will not turn out skewed.

### determineParticleSizes.py: ###

  - Analyzes particle images by applying a various amount of filters such as grayscale, CLAHE, thresholding, and sobel, to pinpoint
    particle locations and subsequently calculate certain useful measurements
  - Analysis results are exported and saved into an Excel file of the name of your choosing that includes a barebones summary sheet that you may add on to
  - Particles are assumed to be elliptical, so the majority of measurements are derived from the major and minor axes of a bounding box except for manually counted pixel area
  - Average Height for calculations such as sphericity, surface area, and volume, can be optionally input by the user at runtime. Otherwise, it is automatically
    calculated by averaging the major and minor axes.
  - Resulting images are also saved in a named test results folder that have bounding rectangles and contours directly drawn onto particles for ease of viewing

### repeatParticleRemoval.py: ###

  - After findcontours.py is run, call this script to generate a new excel results file that filters out potential repeated particles so that results are
    less skewed. The original excel file will remain unchanged.
  - The new excel file will contain 3 sheets: filtered_data, original_data, and summary (based on the filtered_data)
  - Warns the user if more than a certain percentage of particles, designated by MAX_PERCENT_REMOVED, are filtered out

### testing.py: ###

  - Runs regression tests on the images within folder "Test Images Sets"

## INSTALLATION/SETUP ##

- Install Python 3 in 64-bit (**DO NOT DOWNLOAD THE FIRST LINK ON THE HOMEPAGE OF https://www.python.org/. That is 32-bit Python**)
- Set up the environment variable on your computer of choice so you can easily type in 'python' to get the interpreter running rather than the entire path name (Windows Tutorial: https://www.youtube.com/watch?v=Y2q_b4ugPWk)
- pip should be installed by default, so type in the below installs to get the proper compatible versions of these modules for your version of Python
```
    python -m pip install opencv-python
    python -m pip install numpy
    python -m pip install xlsxwriter
    python -m pip install pandas
    python -m pip install openpyxl
```

- To test that everything is installed correctly, type into the interpreter:
```
    import cv2
    import numpy as py
    import xlsxwriter
    import pandas as pd
    import openpyxl
    print(cv2.__version__)
    print(py.__version__)
    print(xlsxwriter.__version__)
    print(pd.__version__)
    print(openpyxl.__version__)
```

Once version numbers can be printed back, the dependencies are set up correctly!

## HOW TO USE IT ##

### RUNNING determineParticleSizes.py ###

- Place all image files in a named folder that is reachable with a path
- In the section under #PLEASE MODIFY#, feel free to modify any of these values:
    - **IMAGE_FOLDER_PATH ->** the path of the folder in which you have your test images (modify the parameter within the call to pathlib.Path())
    - **RESULTS_FILENAME ->** the full name you would like the resulting excel file will be saved as (modify the parameter within the call to pathlib.Path())
    - **TEST_RESULTS_PATH ->** the path and name of the folder the resulting test images will be saved under (modify the parameter within the call to pathlib.Path())
    - **CUSTOM_THRESH ->** turn False if you would like the Otsu thresholding algorithm applied on the images
                               turn True to apply your custom threshold value
    - **THRESH_PARAM ->** used in conjunction with a True **CUSTOM_THRESH**. Enter a number between 0-255, where smaller numbers suggest higher contrast but possible
                               loss of information
    - **TEST ->** testing toggle. Set True if you would like images written out at each step of the analysis process
- Now you are ready to run the code!
    - Change into the correct directory (FSI_Python)
    - Type this command into the command line: `python determineParticleSizes.py`

### RUNNING repeatParticleRemoval.py ###

- **DO NOT RUN THIS SCRIPT UNLESS YOU HAVE A PRE-EXISTING RESULT EXCEL SHEET PRODUCED BY determineParticleSizes.py!!!**
- Modify the global constants at the top as needed:
    - **ORIGINAL_XL_FILENAME ->**     the full file path of the excel file you would like to filter (modify the parameter within the call to pathlib.Path())
    - **NEW_XL_FILENAME ->**          the full file path of the new excel file that will have the filtered results (modify the parameter within the call to pathlib.Path())
    - **POSITIONAL_REMOVAL_RANGE ->** the amount of difference you will allow the x and y coordinates of particles to have to be considered the same
    - **AREA_REMOVAL_RANGE ->**       the amount of difference you will allow the area of particles to have to be considered the same
    - **MAX_PERCENT_REMOVED ->**       if percentage above this number is removed by this program, an error message or suggestion to clean the flowcell will pop up to
                                      inform the user that too many particles or bubbles are stuck to the surface
- Now you are ready to run the code!
    - Change into the correct directory (FSI_Python)
    - Type this command into the command line: `python repeatParticleRemoval.py`

## TESTING ##

This section is dedicated to those who would like to modify, improve, and/or extend this code. It hopes to provide a baseline for testing expected results using manually produced test images that contain uniform or predictable particle sizes.

**[Descriptions of Individual Test Images can be found here](https://github.com/mysticalflyte/fsi_python_image_analysis/wiki/4.-Testing-Sets-Image-Descriptions)**

** NOTE: All tests from Test Sets 1-3 should pass in theory, but one currently fails due to the current program being unable to process vertical 1-pixel-width particles because they are considered too thin during Sobel filtering. In the grand scheme of things, we hope that a large particle count as well as different particle orientations will make up for this elimination.
None of Set 4 passes currently within the designated 5% error - ideally we must create more test images with the particle sets in SolidWorks so that the results average out. Otherwise, we should try to figure out a better way to estimate particle height besides by averaging 2D dimensions with each other. **

### RUNNING testing.py ###

- Modify the global constants at the top as needed:
    - **ALLOWED_RANGE_DIFF ->**   the number of pixels of difference you will allow between the results and your expected values (for sharp test images)
    - **BLURRY_RANGE_DIFF ->**    the number of pixels of difference you will allow between the results and your expected values (for specifically blurry test images)
    - **PERCENT_ERROR ->**  Applicable to only the images for Test Set 4 (the set of mockup 3D images created in SolidWorks), this determines the range of difference you will allow for surface areas and volumes
- Now you are ready to run the code!
    - Change into the correct directory (FSI_Python)
    - Type this command into the command line: `python testing.py`
