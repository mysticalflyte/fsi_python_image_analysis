# Python 3.6.5 opencv2 script for post-image processing
# Specifically exports areas, diameters, and other units of measure of particles
# during floc strength testing and finds averages, minimums, and maximums.
# Written by Jenny Liu

# Imports:
import cv2
import numpy as np
import xlsxwriter as xls
import math
from itertools import compress
import pathlib
import os, os.path

# Global declarations:
############################## DO NOT MODIFY ###################################
PROJECTED_PIXEL_SIZE = 3.45
CLARITY_THRESHOLD = 10
AREA_THRESHOLD_MIN = 3
AREA_THRESHOLD_MAX = 18000

# A list of cropped images to be written out when TEST is True
crops = []

# Global Lists of Measurements:
filtered_min_area_rects = []
auto_areas = []
pixel_areas = []
auto_diameters = []
pixel_diameters = []
major_axes = []
minor_axes = []
aspect_ratios = []
eccentricities = []
surface_areas = []
sauter_diameters = []
volumes = []
sphericities = []


# Modified with user prompt - PLEASE DO NOT CHANGE THIS HERE
AVG_PARTICLE_HEIGHT = -1.0

############################# PLEASE MODIFY  ###################################
IMAGE_FOLDER_PATH = str(pathlib.Path("../Test Images/First Sample Images"))
TEST_RESULTS_PATH = str(pathlib.Path(IMAGE_FOLDER_PATH + "/img_results"))
RESULTS_FILENAME = str(pathlib.Path(TEST_RESULTS_PATH + "/results_test.xlsx"))


# Turn True if you would like to customize the threshold parameter!
# False results in the default Otsu Algorithm optimum threshold calculation.
CUSTOM_THRESH = True
THRESH_PARAM = 80

# Testing toggle. If True, writes out each step to image files.
TEST = True

################################## MAIN CODE  #####################################
def main():
    # Make sure that the optimized version of the code in cv2 is used here
    cv2.setUseOptimized(True)

    # Create test image folder
    pathlib.Path(TEST_RESULTS_PATH).mkdir(exist_ok = True)
    pathlib.Path(TEST_RESULTS_PATH + "/crops").mkdir(exist_ok = True)

    # Set up the excel file to be modified and set the row in the excel file
    # to start writing at
    workbook, xl_sheet_data, xl_sheet_summary = setup_xl_file()
    num_particles = 0
    startRow = 1

    # Optionally have user input an estimate for particle height
    request_height()

    # Make a list of file names in the directory to test and sort them
    file_list = [x for x in os.listdir(IMAGE_FOLDER_PATH) if not os.path.isdir(os.path.join(IMAGE_FOLDER_PATH, x))]

    # Test all images within the folder designated by IMAGE_FOLDER_PATH
    # for file_name in os.listdir(IMAGE_FOLDER_PATH):
    for file_name in file_list:
        # Skip over any non-image files
        if not file_name.endswith(".bmp"):
            continue

        print(file_name)
        img = cv2.imread(os.path.join(IMAGE_FOLDER_PATH, file_name))
        num_particles = analyse(img, startRow, os.path.splitext(file_name)[0], xl_sheet_data)
        startRow = startRow + num_particles
        clear_lists()

    # Write out the summary sheet in the excel workbook, and clear out crops list
    write_xl_summaries(startRow, workbook, xl_sheet_data, xl_sheet_summary)
    crops.clear()
    workbook.close()


# General analysing function. Returns the number of particles successfully
# analysed.
def analyse(img, startRow, file_name, xl_sheet_data):
    test_img(file_name + "_1_original", img)

    img = crop_left_border(img)

    sobel_img, clahe_img = apply_filters(file_name, img)

    thresh_img = threshold_make_binary(clahe_img)

    # Calculate areas and diameters for the particles, both with the contours
    # and by manually counting the pixels
    calc_areas(sobel_img, thresh_img, file_name, img)
    calc_diameters()

    # Find the length of the major and minor axis, the aspect ratios, as well as eccentricity,
    # or how circular the ellipse is (keep in mind, this is merely a cross-section). The closer
    # to 0, the more circular.
    find_side_related_measures()

    # Find the surface areas, volumes, and sauter diameters if the user provides
    # an estimate for the average height of all the particles
    find_height_dependent_measures()

    write_data_to_excel(file_name, startRow, xl_sheet_data)

    return len(auto_areas)


# Apply multiple filters such as grayscale, denoising, clahe, and sobel to the original
# image and return the sobel and clahe results
def apply_filters(file_name, img):
    # Applying a variety of edits and filters to make size calculations easier
    gray_img = grayscale(img)
    test_img(file_name + "_2_gray", gray_img)
    denoise_img = denoise(gray_img)
    test_img(file_name + "_3_denoised", denoise_img)

    # Increase contrast on denoised image and grayscale image using CLAHE
    clahe_denoise_img = increase_contrast(denoise_img)
    clahe_img = increase_contrast(gray_img)
    test_img(file_name + "_4_clahe_denoise", clahe_denoise_img)

    # Used later once bounded rectangles are drawn for determining if the particles
    # are in focus or not.
    sobel_img = sobel_filter(clahe_denoise_img)

    return sobel_img, clahe_img



# Crop the thin left border off the image and then grayscale to analyse further
def crop_left_border(img):
    width = img.shape[1]
    img = img[:,11:]

    return img


# Grayscale the image
def grayscale(img):
    return cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)


# Denoise the image so the particles become more clear
def denoise(img):
    denoise_img = cv2.fastNlMeansDenoising(img, 7, 7, 7)

    return denoise_img


# Use CLAHE (Contrast Limited Adaptive Histogram Equalization) to increase the
# image's contrast.
def increase_contrast(img):
    clahe = cv2.createCLAHE(clipLimit=2.0,)
    clahe_img = clahe.apply(img)

    return clahe_img


# Utilize the Sobel filter, a joint Gaussian smoothing and differentiation
# operation, to make the image more resistant to noise.
# Helps determine whether the particle is in focus or not.
def sobel_filter(img):
    # Apply the filter both horizontally and vertically
    sobelX = cv2.Sobel(img, cv2.CV_32F, dx = 1, dy = 0, ksize = 3, scale = 0.25,
             delta = 0, borderType = cv2.BORDER_DEFAULT)
    sobelY = cv2.Sobel(sobelX, cv2.CV_32F, dx = 0, dy = 1, ksize = 3, scale = 0.25,
             delta = 0, borderType = cv2.BORDER_DEFAULT)
    sobel_img = cv2.magnitude(sobelX, sobelY)
    sobel_img = sobel_img.astype('uint8')

    return sobel_img


# Simple thresholding, basically making the image binary
def threshold_make_binary(img):
    if CUSTOM_THRESH:
        retval, thresh_img = cv2.threshold(img, THRESH_PARAM, 255, cv2.THRESH_BINARY_INV)
    else:
        # Utilize the automatic optimum Otsu Algorithm
        retval, thresh_img = cv2.threshold(img, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

    return thresh_img


# Draws bounded rectangles around particles of interest and ignores unnecessary
# content we don't want to analyse in the image, such as particles that are
# too transparent or too close to the edge.
def calc_areas(sobel_img, thresh_img, file_name, img):
    # Plan to modify this global variable
    global filtered_min_area_rects

    # Used to later manually pixel count
    thresh_rgb_img = cv2.cvtColor(thresh_img, cv2.COLOR_GRAY2BGR)

    __, contours, hierarchy = cv2.findContours(thresh_img, cv2.RETR_EXTERNAL,
                              cv2.CHAIN_APPROX_SIMPLE);

    # Create a list of bounding rectangles and rotated rectangles to later
    # filter and analyse for area
    all_bound_rects = []
    min_area_rects = []
    for contour in contours:
        all_bound_rects.append(cv2.boundingRect(contour))
        min_area_rects.append(cv2.minAreaRect(contour))

    print("Total Number of Contours (Pre-Elimination) = " + str(len(contours)))

    # Dimensions for future reference/calculations
    height, width, channels = thresh_rgb_img.shape
    xMax = width - 2
    yMax = height - 2

    # Used to filter out unnecessary bounded rectangles
    toKeep = []

    for i in range(len(min_area_rects)):
        #Crop out the area of interest with the individual particle to examine
        center, size, angle = min_area_rects[i]

        # Crop out the rectangles from the sobel image as well as the threshold image
        x, y, w, h = all_bound_rects[i]
        sobel_roi_crop = sobel_img[y : y + h, x: x + w]
        threshold_roi_crop = thresh_img[y: y + h, x : x + w]

        # Find the minimum and maximum of this current particle's ROI
        (minval, maxval, minloc, maxloc) = cv2.minMaxLoc(sobel_roi_crop)

        # Calculate manual-area by counting number of non-black (white) pixels
        num_white_pixels = count_white_pixels(threshold_roi_crop)

        # Determine whether or not to keep the particle we just cropped
        kept = acceptable_particle(num_white_pixels, all_bound_rects[i],
            xMax, yMax, maxval)
        toKeep.append(kept)

        if (kept):
            # Calculate auto-generated area with contours
            contour_area = cv2.contourArea(contours[i])
            auto_areas.append(contour_area)
            pixel_areas.append(num_white_pixels)

            crops.append(threshold_roi_crop)
            # test_img("crops/"+file_name + "_crop_" + str(i), threshold_roi_crop)


    print("Total Number of Contours (Post-Elimination): " + str(len(auto_areas)))

    filtered_min_area_rects = list(compress(min_area_rects, toKeep))

    save_thresh_roi_crops()
    draw_rect_img(thresh_rgb_img, img, contours, file_name)


# Manually count the number of white pixels that are present in the threshold_roi_crop
# from draw_bounded_rects
def count_white_pixels(threshold_roi_crop):
    num = 0
    height, width = threshold_roi_crop.shape
    for y in range(height):
        for x in range(width):
            if threshold_roi_crop[y][x] != 0:
                num += 1

    return num


# Determines whether the corresponding particle is acceptable for calculation.
# Conditions include not too close to the edge and not too transparent for
# accuracy.
def acceptable_particle(num_white_pixels, bound_rect, xMax, yMax, maxval):
    (x, y, width, height) = bound_rect
    if (maxval > CLARITY_THRESHOLD and num_white_pixels > AREA_THRESHOLD_MIN and
        num_white_pixels < AREA_THRESHOLD_MAX):
        if (x > 1 and y > 1 and (x + width <= xMax and y + height <= yMax)):
            return True
        else:
            return False
    else:
        return False


# Calculate diameters using the auto generated areas and pixel-counted areas
def calc_diameters():
    for i in range(len(auto_areas)):
        auto_diameter = PROJECTED_PIXEL_SIZE * (math.sqrt(4 * auto_areas[i] / math.pi));
        pixel_diameter = PROJECTED_PIXEL_SIZE * (math.sqrt(4 * pixel_areas[i] / math.pi))
        auto_diameters.append(auto_diameter)
        pixel_diameters.append(pixel_diameter)


# Save the major and minor axes properly, then derive the eccentricity as well as
# aspect ratio from them
def find_side_related_measures():
    for min_area_rect in filtered_min_area_rects:
        center, size, angle = min_area_rect
        x, y = center
        side1, side2 = size
        minor_axis = side1
        major_axis = side2
        if major_axis < minor_axis:
            minor_axis = side2
            major_axis = side1
        major_axes.append(major_axis)
        minor_axes.append(minor_axis)

        # Derive the eccentricity
        find_cross_section_eccentricity(minor_axis, major_axis)

        # Derive the aspect aspect ratio
        find_aspect_ratio(minor_axis, major_axis)


# Calculate the eccentricities, or how circular the ellipses contained in the
# minimum area rectangles are.
def find_cross_section_eccentricity(minor_axis, major_axis):
    squared_e = 1 - (math.pow(minor_axis, 2) / math.pow(major_axis, 2))
    e = math.sqrt(squared_e)
    eccentricities.append(e)


# Calculate the aspect ratio, or the ratio of the minor_axis to the major_axis
def find_aspect_ratio(minor_axis, major_axis):
    ratio = minor_axis / major_axis
    aspect_ratios.append(ratio)


# Set particle height for further measurements if the user defines a value at
# the start.
def request_height():
    global AVG_PARTICLE_HEIGHT
    do_calc = input("Do you have an estimate for average particle height? Y/N ")
    if (do_calc == 'y' or do_calc == 'Y'):
        # Make sure the user inputs a valid number!
        while True:
            try:
                AVG_PARTICLE_HEIGHT = float(input("Please enter your average particle height: "))
                if isinstance(AVG_PARTICLE_HEIGHT, float):
                    break
            except:
                pass
            print('\nInvalid input, please try again.')


# If height is not provided, estimate height based on the width and length of the
# particle's bounding box. You may change this function if you would like a different
# way to estimate.
# Currently, returns the average of the major axis and minor axis lengths as the particle's height.
def estimate_height(major_axis, minor_axis):
    return (major_axis + minor_axis) / 2


# Depending on user input, calculate the surface area, volume, and sauter diameter
# of all particles assuming they have a similar average height and are shaped like
# ellipsoids
def find_height_dependent_measures():
    # Set c in whatever the user input
    if AVG_PARTICLE_HEIGHT != -1:
        c = AVG_PARTICLE_HEIGHT / 2
    current_surface_area = 0

    # Iterate through all min area rects to calculate height-dependent measures
    for i in range(len(filtered_min_area_rects)):
        center, size, angle = filtered_min_area_rects[i]
        major_axis, minor_axis = size
        a = major_axis / 2
        b = minor_axis / 2

        # Height wasn't set by user, so set c based on an estimate
        if AVG_PARTICLE_HEIGHT == -1:
            c = estimate_height(major_axis, minor_axis) / 2

        # Calculate heigh-dependent measurements now!
        current_surface_area = calc_ellipsoid_surface_area(a, b, c)
        current_volume = calc_ellipsoid_volume(a, b, c)
        calc_sauter_diameter(current_surface_area)

        calc_sphericity(current_surface_area, current_volume)


# Surface Area calculation
def calc_ellipsoid_surface_area(a, b, c):
    numerator = math.pow(a * b, 1.6) + math.pow(a * c, 1.6) + math.pow(b * c, 1.6)
    surface_area = 4 * math.pi * math.pow(numerator / 3, 1/1.6)
    surface_areas.append(surface_area)

    return surface_area


# Volume calculation
def calc_ellipsoid_volume(a, b, c):
    volume = (4 / 3) * math.pi * a * b * c
    volumes.append(volume)

    return volume


# If surface areas and volumes were calculated (as in the user gave a potential
# height input), then move on to find sauter diameters.
def calc_sauter_diameter(current_surface_area):
    surface_diameter = math.sqrt(current_surface_area / math.pi)
    sauter_diameters.append(surface_diameter)


# Calculate sphericity only if height has been provided and surface area/volume
# has also been estimated
def calc_sphericity(current_surface_area, current_volume):
    radius_for_volume = math.pow(current_volume * (3/4) / math.pi, 1/3)
    sphere_surface_area = 4 * math.pi * math.pow(radius_for_volume, 2)
    sphericity = sphere_surface_area / current_surface_area
    sphericities.append(sphericity)


# Clean up lists for the next image to populate
def clear_lists():
    filtered_min_area_rects.clear()
    auto_areas.clear()
    pixel_areas.clear()
    auto_diameters.clear()
    pixel_diameters.clear()
    eccentricities.clear()
    aspect_ratios.clear()
    minor_axes.clear()
    major_axes.clear()

    # clear height-dependent lists
    if (AVG_PARTICLE_HEIGHT != -1):
        surface_areas.clear()
        sauter_diameters.clear()
        volumes.clear()
        sphericities.clear()


# Create the excel sheet to be edited. All data will end up in such a file called
# 'results.xlsx'
def setup_xl_file():
    # Create excel file
    xl_filename = RESULTS_FILENAME
    workbook = xls.Workbook(xl_filename)
    xl_sheet_data = workbook.add_worksheet("data")
    xl_sheet_summary = workbook.add_worksheet("summary")
    bold = workbook.add_format({'bold': 1})
    xl_sheet_data.set_column(0, 14, 15)
    xl_sheet_summary.set_column(0, 1, 25)

    #Write column headers for data
    xl_sheet_data.write('A1', 'File_Name', bold)
    xl_sheet_data.write('B1', 'Pixel_Area', bold)
    xl_sheet_data.write('C1', 'Pixel_Diameter', bold)
    xl_sheet_data.write('D1', 'Contour_Area', bold)
    xl_sheet_data.write('E1', 'Contour_Diameter', bold)
    xl_sheet_data.write('F1', 'Major_axis', bold)
    xl_sheet_data.write('G1', 'Minor_axis', bold)
    xl_sheet_data.write('H1', 'Aspect_Ratio', bold)
    xl_sheet_data.write('I1', 'Eccentricity', bold)
    xl_sheet_data.write('J1', 'Surface_Area', bold)
    xl_sheet_data.write('K1', 'Sauter_Diameter', bold)
    xl_sheet_data.write('L1', 'Volume', bold)
    xl_sheet_data.write('M1', 'Sphericity', bold)
    xl_sheet_data.write('N1', 'X_coord', bold)
    xl_sheet_data.write('O1', 'Y_coord', bold)

    return workbook, xl_sheet_data, xl_sheet_summary


# Write out data to the excel file
def write_data_to_excel(filename, startRow, xl_sheet_data):
    row = 0
    for i in range(len(filtered_min_area_rects)):
        row = startRow + i
        center, size, angle = filtered_min_area_rects[i]
        x, y = center
        xl_sheet_data.write(row, 0, filename)
        xl_sheet_data.write_number(row, 1, pixel_areas[i])
        xl_sheet_data.write_number(row, 2, pixel_diameters[i])
        xl_sheet_data.write_number(row, 3, auto_areas[i])
        xl_sheet_data.write_number(row, 4, auto_diameters[i])
        xl_sheet_data.write_number(row, 5, major_axes[i])
        xl_sheet_data.write_number(row, 6, minor_axes[i])
        xl_sheet_data.write_number(row, 7, aspect_ratios[i])
        xl_sheet_data.write_number(row, 8, eccentricities[i])

        # Write height-dependent calculations out, depending on user input
        xl_sheet_data.write_number(row, 9, surface_areas[i])
        xl_sheet_data.write(row, 10, sauter_diameters[i])
        xl_sheet_data.write_number(row, 11, volumes[i])
        xl_sheet_data.write_number(row, 12, sphericities[i])

        xl_sheet_data.write_number(row, 13, x)
        xl_sheet_data.write_number(row, 14, y)


    return row


# Write the average functions for the excel sheet to see the overall area and
# diameter averages
def write_xl_summaries(numData, workbook, xl_sheet1, xl_sheet_summary):
    numData = str(numData)
    # Write summary section headers
    bold = workbook.add_format({'bold': 1})
    xl_sheet_summary.write('A1', 'SUMMARY', bold)
    xl_sheet_summary.write('A2', 'AVG_PIXEL_AREA:', bold)
    xl_sheet_summary.write('A3', 'AVG_PIXEL_DIAMETER:', bold)
    xl_sheet_summary.write('A4', 'AVG_CONTOUR_AREA:', bold)
    xl_sheet_summary.write('A5', 'AVG_CONTOUR_DIAMETER:', bold)
    xl_sheet_summary.write('A6', 'AVG_MINOR_AXIS:', bold)
    xl_sheet_summary.write('A7', 'AVG_MAJOR_AXIS:', bold)
    xl_sheet_summary.write('A8', 'AVG_ASPECT_RATIO:', bold)
    xl_sheet_summary.write('A9', 'AVG_ECCENTRICITY:', bold)
    xl_sheet_summary.write('A10', 'MAX_ECCENTRICITY:', bold)
    xl_sheet_summary.write('A11', 'MIN_PIXEL_AREA:', bold)
    xl_sheet_summary.write('A12', 'MAX_PIXEL_AREA:', bold)
    xl_sheet_summary.write('A13', 'SAUTER_MEAN_DIAMETER:', bold)
    xl_sheet_summary.write('A14', 'AVG_VOLUME:', bold)
    xl_sheet_summary.write('A15', 'AVG_SPHERICITY:', bold)
    xl_sheet_summary.write('A16', 'AVG_SURFACE_AREA:', bold)



    # Write actual summary functions
    xl_sheet_summary.write(1, 1, "=AVERAGE(" + xl_sheet1.get_name() + "!B2:" + xl_sheet1.get_name() + "!B" + numData + ")")
    xl_sheet_summary.write(2, 1, "=AVERAGE(" + xl_sheet1.get_name() + "!C2:" + xl_sheet1.get_name() + "!C" + numData + ")")
    xl_sheet_summary.write(3, 1, "=AVERAGE(" + xl_sheet1.get_name() + "!D2:" + xl_sheet1.get_name() + "!D" + numData + ")")
    xl_sheet_summary.write(4, 1, "=AVERAGE(" + xl_sheet1.get_name() + "!E2:" + xl_sheet1.get_name() + "!E" + numData + ")")
    xl_sheet_summary.write(5, 1, "=AVERAGE(" + xl_sheet1.get_name() + "!F2:" + xl_sheet1.get_name() + "!F" + numData + ")")
    xl_sheet_summary.write(6, 1, "=AVERAGE(" + xl_sheet1.get_name() + "!G2:" + xl_sheet1.get_name() + "!G" + numData + ")")
    xl_sheet_summary.write(7, 1, "=AVERAGE(" + xl_sheet1.get_name() + "!H2:" + xl_sheet1.get_name() + "!H" + numData + ")")
    xl_sheet_summary.write(8, 1, "=AVERAGE(" + xl_sheet1.get_name() + "!I2:" + xl_sheet1.get_name() + "!I" + numData + ")")
    xl_sheet_summary.write(9, 1, "=MAX(" + xl_sheet1.get_name() + "!I2:" + xl_sheet1.get_name() + "!I" + numData + ")")
    xl_sheet_summary.write(10, 1, "=MIN(" + xl_sheet1.get_name() + "!B2:" + xl_sheet1.get_name() + "!B" + numData + ")")
    xl_sheet_summary.write(11, 1, "=MAX(" + xl_sheet1.get_name() + "!B2:" + xl_sheet1.get_name() + "!B" + numData + ")")
    xl_sheet_summary.write(12, 1, "=AVERAGE(" + xl_sheet1.get_name() + "!K2:" + xl_sheet1.get_name() + "!K" + numData + ")")
    xl_sheet_summary.write(13, 1, "=AVERAGE(" + xl_sheet1.get_name() + "!L2:" + xl_sheet1.get_name() + "!L" + numData + ")")
    xl_sheet_summary.write(14, 1, "=AVERAGE(" + xl_sheet1.get_name() + "!M2:" + xl_sheet1.get_name() + "!M" + numData + ")")
    xl_sheet_summary.write(15, 1, "=AVERAGE(" + xl_sheet1.get_name() + "!J2:" + xl_sheet1.get_name() + "!J" + numData + ")")


# For testing purposes. Draws the bounding rectangles around the particles that
# were analysed and writes this out to an image file called "rect_threshold_image"
def draw_rect_img(thresh_rgb_img, img, contours, file_name):
    for i in range(len(filtered_min_area_rects)):
        # Draw rotated best-fit rectangles
        rotated_rect = cv2.boxPoints(filtered_min_area_rects[i])
        rotated_rect = np.int0(rotated_rect)
        cv2.drawContours(thresh_rgb_img, [rotated_rect], -1, (0,255,0), 2)
        cv2.drawContours(img, [rotated_rect], -1, (0,0,255), 2)
        cv2.drawContours(img, contours, -1, (0,0,255), 3)

    cv2.imwrite(str(pathlib.Path(TEST_RESULTS_PATH + "/" + file_name + "_5_rect_thresh_image.bmp")), thresh_rgb_img)
    cv2.imwrite(str(pathlib.Path(TEST_RESULTS_PATH + "/" + file_name + "_6_rect_og_image.bmp")), img)


# For testing purposes, save all cropped threshold images that contain reasonable particles into a folder
# called "crops"
def save_thresh_roi_crops():
    if TEST:
        for i in range(len(crops)):
            cv2.imwrite(str(pathlib.Path(TEST_RESULTS_PATH + "/crops/crop" + str(i) + ".bmp")), crops[i])


# For the purpose of seeing what is happening to the images at each step.
# If TEST = True, writes out the image files into a folder designated by
# TEST_RESULTS_PATH
def test_img(name, img):
    if TEST:
        file_path = str(pathlib.Path(TEST_RESULTS_PATH + "/" + name + ".bmp"))
        cv2.imwrite(file_path, img)


# Run the main program
if __name__ == "__main__":
    main()
