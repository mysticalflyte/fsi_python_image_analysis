# Python 3.6.5 Testing script to check determineParticleSizes result validity
# Uses the simple testing set
# Written by Jenny Liu

# Imports:
import determineParticleSizes
import cv2
import pathlib
import numpy as np
import pandas as pd
import openpyxl


# Constants:
TEST_FOLDER = "TestSets"
SET1_FOLDER = "Set1"
SET2_FOLDER = "Set2"
SET3_FOLDER = "Set3"
SET4_FOLDER = "Set4"

IMG_RESULTS_FOLDER = "img_results"

# Customizable Values:
ALLOWED_RANGE_DIFF = 5
BLURRY_RANGE_DIFF = 100
PERCENT_ERROR = 0.05

# Summary of Outcomes:
NUM_PASS = 0
NUM_FAIL = 0


def main():
    # Set the projected pixel size so there is no scale factor to be multiplied during calculation
    determineParticleSizes.PROJECTED_PIXEL_SIZE = 1

    # Test the 4 sets of test images
    test_set1()
    test_set2()
    test_set3()
    test_set4()

    # Print out the total number of passed and failed tests!
    print_summary()


# Runs determineParticleSizes analysis code on a particular test image and verifies its results
# based on an expected particle count, pixel area, and pixel diameter
def test_individual_uniform_img(set_name, file_num, expect_count, expect_area, expect_diameter):
    test_img, file_path = get_test_img(file_num, set_name)

    # Setup individual excel file for this current image
    determineParticleSizes.RESULTS_FILENAME = file_path + "_results.xlsx"
    workbook, xl_sheet_data, xl_sheet_summary = determineParticleSizes.setup_xl_file()

    # Test if particle count is correct
    startRow = 1
    num_particles = determineParticleSizes.analyse(test_img, startRow, file_num, xl_sheet_data)
    print_result(file_path, num_particles == expect_count, "particle count", expect_count, num_particles)

    # Test if pixel_areas are uniformly correct
    print_result(file_path, verify_list(determineParticleSizes.pixel_areas, expect_area, None),
        "pixel area", expect_area, determineParticleSizes.pixel_areas)

    # Test if pixel_diameters are uniformly correct
    print_result(file_path, verify_list(determineParticleSizes.pixel_diameters, expect_diameter, None),
        "pixel_diameter", expect_diameter, determineParticleSizes.pixel_diameters)
    print("\n")

    # Draw the bounding rectangles and contours on the image for ease of viewing
    test_img = determineParticleSizes.crop_left_border(test_img)
    draw_contours_and_rects(test_img, file_num, set_name)

    # Write out summary of averages, mins, and maxes out to the excel file
    determineParticleSizes.write_xl_summaries(startRow, workbook, xl_sheet_data, xl_sheet_summary)

    workbook.close()
    determineParticleSizes.clear_lists()


# Gather the test image to analyse
def get_test_img(file_num, set_name):
    file_name = set_name + "_Image" + str(file_num)
    file_type = ".bmp"
    file_path = str(pathlib.Path(TEST_FOLDER + "/" + set_name + "/" + file_name))

    test_img = cv2.imread(file_path + file_type)

    return test_img, file_path


# Verify all the values of the passed in list against the expected value
# Return true if they all match as expected, and false otherwise
def verify_list(list, expect_val, expect_val_list):
    for i in range(len(list)):
        val = list[i]
        if expect_val_list == None:
            expect_val = expect_val
        else:
            expect_val = expect_val_list[i]
        if val < expect_val - ALLOWED_RANGE_DIFF or val > expect_val + ALLOWED_RANGE_DIFF:
            return False
    return True


# Print out the result of this test!
def print_result(file_path, condition, test_type, expect_val, actual_val):
    global NUM_FAIL
    global NUM_PASS
    if (condition):
        print(file_path + " " + test_type + " PASSED!")
        NUM_PASS = NUM_PASS + 1
    else:
        print(file_path + " " + test_type + " FAILED!")
        print("Expected " + test_type + ": " + str(expect_val))
        NUM_FAIL = NUM_FAIL + 1


# Draws and saves a version of the current test image with the bounding min area rectangles
# drawn on top
def draw_contours_and_rects(test_img, file_num, set_name):
    for i in range(len(determineParticleSizes.filtered_min_area_rects)):
        # Draw rotated best-fit rectangles
        rotated_rect = cv2.boxPoints(determineParticleSizes.filtered_min_area_rects[i])
        rotated_rect = np.int0(rotated_rect)
        cv2.drawContours(test_img, [rotated_rect], -1, (0,0,255), 2)

    results_path = str(pathlib.Path(TEST_FOLDER + set_name + "/" + IMG_RESULTS_FOLDER + "/rect_og_image_" + str(file_num) + ".bmp"))
    cv2.imwrite(results_path, test_img)


# Tests all images in test set1
# Parameters for individual image tests in order: set_name, file_num, expect_count, expect_area, expect_diameter
def test_set1():
    global ALLOWED_RANGE_DIFF

    # Create test image folder
    pathlib.Path(TEST_FOLDER + "/" + SET1_FOLDER + "/" + IMG_RESULTS_FOLDER).mkdir(exist_ok = True)

    # Blurry particles - can't estimate pixel area definitely so increase range
    ALLOWED_RANGE_DIFF = ALLOWED_RANGE_DIFF + BLURRY_RANGE_DIFF
    test_individual_uniform_img(SET1_FOLDER, 1, 100, 32, 6)

    # Clear hard-edged particle images, so return range to normal
    ALLOWED_RANGE_DIFF = ALLOWED_RANGE_DIFF - BLURRY_RANGE_DIFF
    test_individual_uniform_img(SET1_FOLDER, 2, 100, 32, 6)
    test_individual_uniform_img(SET1_FOLDER, 3, 1200, 32, 6)
    test_individual_uniform_img(SET1_FOLDER, 4, 100, 32, 6)
    test_individual_uniform_img(SET1_FOLDER, 5, 100, 32, 6)

    # Too low contrast - currently determineParticleSizes finds none of the 100 particles!
    test_individual_uniform_img(SET1_FOLDER, 6, 0, 32, 6)

    # Non-uniform particles
    # NOTE: Can't register the vertical particles, maybe filtered out by sobel?
    test_individual_uniform_img(SET1_FOLDER, 7, 100, 32, 6)


# Tests all images in test set2
# Parameters for individual image tests in order: set_name, file_num, expect_count, expect_area, expect_diameter
def test_set2():
    global ALLOWED_RANGE_DIFF

    # Create test image folder
    pathlib.Path(TEST_FOLDER + "/" + SET2_FOLDER + "/" + IMG_RESULTS_FOLDER).mkdir(exist_ok = True)

    test_individual_uniform_img(SET2_FOLDER, 1, 10, 4311, 75)
    test_individual_uniform_img(SET2_FOLDER, 2, 100, 4311, 75)
    test_individual_uniform_img(SET2_FOLDER, 3, 20, 4311, 75)
    test_individual_uniform_img(SET2_FOLDER, 4, 20, 4311, 75)
    test_individual_uniform_img(SET2_FOLDER, 5, 20, 4311, 75)

    # Blurry particles - can't estimate pixel area definitely so increase range
    ALLOWED_RANGE_DIFF = ALLOWED_RANGE_DIFF + BLURRY_RANGE_DIFF
    test_individual_uniform_img(SET2_FOLDER, 6, 20, 4311, 75)

    # Non-uniform particles of estimated 945 pixel area and 35 pixel diameter
    ALLOWED_RANGE_DIFF = ALLOWED_RANGE_DIFF - BLURRY_RANGE_DIFF
    test_individual_uniform_img(SET2_FOLDER, 7, 20, 945, 35)
    # test_individual_uniform_img(SET2_FOLDER, 7, 20, 945, 75)



# Tests all images in test set3
def test_set3():
    # Create test image folder
    pathlib.Path(TEST_FOLDER + "/" + SET3_FOLDER + "/" + IMG_RESULTS_FOLDER).mkdir(exist_ok = True)

    file_num = 1
    test_img, file_path = get_test_img(file_num, SET3_FOLDER)

    # Setup individual excel file for this current image
    determineParticleSizes.RESULTS_FILENAME = file_path + "_results.xlsx"
    workbook, xl_sheet_data, xl_sheet_summary = determineParticleSizes.setup_xl_file()

    # Test if particle count is correct
    startRow = 1
    num_particles = determineParticleSizes.analyse(test_img, startRow, file_num, xl_sheet_data)

    # Check the amount of particles captured is correct
    expect_count = 10
    print_result(file_path, num_particles == expect_count, "particle count", expect_count, num_particles)

    # Check the areas against this list
    expected_areas = [400, 225, 100, 81, 64, 49, 36, 16, 9, 4]
    print_result(file_path, verify_list(determineParticleSizes.pixel_areas, None, expected_areas),
        "pixel area", expected_areas, determineParticleSizes.pixel_areas)

    # Test if pixel_diameters are uniformly correct
    expected_diameters = [20, 15, 10, 9, 8, 7, 6, 4, 3, 2]
    print_result(file_path, verify_list(determineParticleSizes.pixel_diameters, None, expected_diameters),
        "pixel_diameter", expected_diameters, determineParticleSizes.pixel_diameters)
    print("\n")

    draw_contours_and_rects(test_img, file_num, SET3_FOLDER)

    determineParticleSizes.write_xl_summaries(startRow, workbook, xl_sheet_data, xl_sheet_summary)
    workbook.close()


# Test one set of 3D images for surface area and volume
def test_3D_set(shape_type, expect_surf_area, expect_vol):
    # Create test image folder and setup constants
    determineParticleSizes.IMAGE_FOLDER_PATH = TEST_FOLDER + "/" + SET4_FOLDER + "/"
    pathlib.Path(determineParticleSizes.IMAGE_FOLDER_PATH + IMG_RESULTS_FOLDER).mkdir(exist_ok = True)
    determineParticleSizes.RESULTS_FILENAME = pathlib.Path(determineParticleSizes.IMAGE_FOLDER_PATH + shape_type + "_results.xlsx")
    workbook, xl_sheet_data, xl_sheet_summary = determineParticleSizes.setup_xl_file()

    startRow = 1
    # Analyse 5 images of the same shape at a time, and have the results compiled in an excel sheet
    for i in range(1, 6):
        # Retrieve file to test
        file_name = shape_type + " (" + str(i) + ").png"
        img = cv2.imread(determineParticleSizes.IMAGE_FOLDER_PATH + file_name)

        # Analyse file
        num_particles = determineParticleSizes.analyse(img, startRow, i, xl_sheet_data)
        startRow = startRow + num_particles
        determineParticleSizes.clear_lists()

    # Write out summary sheet to excel workbook
    determineParticleSizes.write_xl_summaries(startRow, workbook, xl_sheet_data, xl_sheet_summary)
    workbook.close()

    verify_avg_measures(expect_surf_area, expect_vol)


# Open the excel workbook, read in the relevant summary average values and compare them against
# the expected values. Print out results.
def verify_avg_measures(expect_surf_area, expect_vol):
    # Obtain the values in the summary sheet for comparison
    wb = openpyxl.load_workbook(filename=determineParticleSizes.RESULTS_FILENAME, data_only=True)
    sheet = wb["summary"]
    # summary_surf_area = float(sheet['B16'].value)
    # summary_vol = float(sheet['B14'].value)
    summary_surf_area = float(sheet['B16'].internal_value)
    summary_vol = float(sheet['B14'].internal_value)

    # Find range of error
    surf_area_min, surf_area_max = reasonable_error(expect_surf_area)
    vol_min, vol_max = reasonable_error(expect_surf_area)

    #Print out results
    # def print_result(file_path, condition, test_type, expect_val, actual_val):
    surf_area_good = surf_area_min < summary_surf_area and summary_surf_area < surf_area_max
    vol_good = vol_min < summary_vol and summary_vol < vol_max
    print_result(str(determineParticleSizes.RESULTS_FILENAME), surf_area_good, "surface area", expect_surf_area, summary_surf_area)
    print_result(str(determineParticleSizes.RESULTS_FILENAME), vol_good, "volume", expect_vol, summary_vol)


# Return the val - PERCENT_ERROR and val + PERCENT_ERROR as a tuple, basically generating a reasonable range of error to compare
# an actual experiment value against
def reasonable_error(val):
    percentDiff = val * PERCENT_ERROR
    min = val - percentDiff
    max = val + percentDiff

    return min, max


# Tests the 2D images from particles created in SolidWorks of uniform shapes and sizes
# Parameters are in this order: shape_type, expect_surf_area, expect_vol
def test_set4():
    test_3D_set("long cyl", 3210.71, 10159.91)
    test_3D_set("med cyl", 1759.29, 5079.96)
    test_3D_set("short cyl", 2997.08, 11196.64)
    test_3D_set("rice", 4600.03, 24810.73)
    test_3D_set("pyr", 5708.42, 24696)


# Print out the number of tests passed and failed for the user
def print_summary():
    print("Number of tests passed: " + str(NUM_PASS))
    print("Number of tests failed: " + str(NUM_FAIL))


# Run the main program
if __name__ == "__main__":
    main()
