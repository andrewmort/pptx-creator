#!/usr/bin/env python

#
# File: pptx-creator.py
# Author: amort
#
# Dependencies:
#   - python-pptx: Install with "pip install --user python-pptx"
#   - pathlib2: (version 3.x only) Install with "pip install --user pathlib2"
#
# Versions History:
#   Version 0.1 (7-30-2018)
#       Initial version.
#   Version 0.2 (9-12-2018)
#       Change input file to xml format
#       Add support for table generation
#       Extract data from csv and xlsx files
#
# Description: This file takes an xml file with slide definitions and
#   creates a powerpoint presentation. The file can specify the format
#   of each slide, including slide title, text, images, etc. The 
#   powerpoint is created using a template file, which can be specified
#   on the command line.
#
#
# Future Improvements (TODO):
#   - Better defaults
#       * use basic powerpoint template
#   - Specify template definition file
#       * Include mapping between name of slide layout and placeholders 
#           in text file and the indexes used to accesses these elements
#

from pptx import Presentation
import datetime
import os
import re
import argparse
import sys
import xml.etree.ElementTree as ET

try:
    import pathlib
except ImportError:
    import pathlib2 as pathlib


default_template = ""
default_filename = "generate_pptx_{}.pptx".format(
        datetime.datetime.now().strftime("%Y%m%d"))

# Power Point Slide Layout Indexes - Original without picture placeholder
PPTX_LAYOUT_TITLE     = 0
PPTX_LAYOUT_SECTION   = 2
PPTX_LAYOUT_1CONTENT  = 1
PPTX_LAYOUT_2CONTENT  = 3

# Power Point Slide Placeholder Indexes - Original without picture placeholder
PPTX_PLACEHOLDER_TITLE_TITLE          = 0
PPTX_PLACEHOLDER_TITLE_SUBTITLE       = 13
PPTX_PLACEHOLDER_SECTION_TITLE        = 12
PPTX_PLACEHOLDER_SECTION_SUBTITLE     = 1
PPTX_PLACEHOLDER_1CONTENT_TITLE       = 0
PPTX_PLACEHOLDER_1CONTENT_CONTENT0    = 1
PPTX_PLACEHOLDER_2CONTENT_TITLE       = 0
PPTX_PLACEHOLDER_2CONTENT_CONTENT0    = 1
PPTX_PLACEHOLDER_2CONTENT_CONTENT1    = 10

def main():
    global verbose

    # **** Setup argument parser ****
    parser = argparse.ArgumentParser(description="This program uses a pptx "
            "definition file and creates a pptx starting from a pptx template.")
    parser.add_argument("input",  help="pptx xml definition file")
    parser.add_argument("-o", "--output", default=default_filename,
            help="pptx output file")
    parser.add_argument("-t", "--template", default=default_template,
            help="pptx template file")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args()


    # **** Get arguments ****
    verbose           = args.verbose
    filename_input    = args.input
    filename_output   = args.output
    filename_template = args.template


    # **** Get input definition ****

    if verbose:
        print("INFO: Reading " + filename_input + "...")

    # Read xml tree from file
    try:
        xml_tree = ET.parse(filename_input)
    except:
        print(sys.exc_info()[0])
        raise


    # **** Create presentation ****

    # Create presentation using template
    prs = Presentation(filename_template)

    # Parse input lines to create presentation
    make_presentation(prs, xml_tree)

    # Save presetnation as filename_output
    prs.save(filename_output)

# Use lines from input file to create presentation
def make_presentation(prs, xml_tree):
    prs_slides = []
    prs_slides_ref = {}
    slides_layout = []
    defines = {}

    slide_idx = 0

    # Create slides first so we can create links while adding content
    for xml_slide in xml_tree.iter('slide'):
        # Get layout type
        layout = xml_slide.get('layout')
        if layout == "title":
            slides_layout[slide_idx] = PPTX_LAYOUT_TITLE
        elif layout == "section":
            slides_layout[slide_idx] = PPTX_LAYOUT_SECTION
        elif layout == "1content":
            slides_layout[slide_idx] = PPTX_LAYOUT_1CONTENT
        elif layout == "2content":
            slides_layout[slide_idx] = PPTX_LAYOUT_2CONTENT
        else:
            print("Error: Slide {} does not specify a valid layout "
                    "attribute \"{}\"".format(slide_idx+1,layout))
            raise

        # Create the new slide
        prs_layout = prs.slide_layouts[slides_layout[slide_idx]]
        prs_slides[slide_idx] = prs.slides.add_slide(prs_layout)

        # If slide has a label, add it to the references
        if (label = xml_slide.get('label')):
            # Do not allow duplicate references
            if (label in slides_ref):
                print("Error: Label {} on slide {} (layout {}) already "
                        "exists".format(label,slide_idx+1,layout))
                raise

            prs_slides_ref[label] = slides[slide_idx]

        slide_idx++

    # Reset variable for loop
    slide_idx = 0

    # Loop through all xml elements in order
    for xml_child in xml_tree:
        # Slide definitions
        if (xml_child.tag == 'slide'):


        # Variable definitions
        elif (xml_child.tag == 'define'):

        # Invalid elements
        else:
            print("Error: Invalid element \"" + xml_child.tag + "\"")
            raise


def old_make_presentation(prs, lines):
    slide_type = ''
    invalid_images = []
    image_path = ''

    for i in range(0, len(lines)):
        line = lines[i]
        line_idx = i + 1

        # Remove comments and end of lines
        line = re.sub(r'\s*#.*', '', line)
        line = re.sub(r'\n', '', line)

        # Split the line into fields seperated by commas
        fields = re.split(r'\s*,\s*', line)

        # Ignore when there are no fields
        if len(fields) < 1 or fields[0] == '':
            continue

        if verbose:
            print("line {}".format(line_idx))
            print(fields)

        # Create new slide
        if fields[0].strip() == 'slide':
            if image_path == '':
                print("Warning: image_path is not set using \".\"")
                image_path = "."

            if len(fields) < 2:
                print("Error: 'slide' line must contain more than 1 field."
                        "(line {})\n".format(line_idx))
                raise

            # Set layout for new slide
            if fields[1].strip() == "layout":
                if len(fields) < 3:
                    print("Error: 'slide, layout' line must contain more "
                            "than 2 fields. (line {})\n".format(line_idx))
                    raise

                # Get slide layout
                if fields[2].strip() == "title":
                    slide_layout = prs.slide_layouts[PPTX_LAYOUT_TITLE]
                elif fields[2].strip() == "section":
                    slide_layout = prs.slide_layouts[PPTX_LAYOUT_SECTION]
                elif fields[2].strip() == "1content":
                    slide_layout = prs.slide_layouts[PPTX_LAYOUT_1CONTENT]
                elif fields[2].strip() == "2content":
                    slide_layout = prs.slide_layouts[PPTX_LAYOUT_2CONTENT]
                else:
                    print("Error: Invalid field 2 \"" + fields[2] +
                            "\" (line {})\n".format(line_idx))
                    raise

                # Save slide type and add new slide
                slide_type = fields[2].strip()
                slide = prs.slides.add_slide(slide_layout)

            else:
                print("Error: Invalid field 1 \"" + fields[1] + 
                        "\" (line {})\n".format(line_idx))
                raise

        elif re.match("\s*image_path.*", line):
            fields = re.split("\s*=\s*", line)
            if len(fields) < 2:
                print("Error: image_path assignment must contain more than 1 "
                        "field (line {})".format(line_idx))
                raise

            if slide_type != '':
                print ("Error: image_path assignment must come before "
                        "any slides are added! (line {})".format(Line_idx))
                raise

            # Remove any quotes
            image_path = re.sub(r'\"', '', fields[1])

            # Find all image files under image path
            images = list(pathlib.Path(image_path).glob('**/*.png'))

        # Current slide operation
        elif slide_type != '':
            invalid = 0

            # Determine field 0 for a placeholder line
            if slide_type == "title":
                if fields[0].strip() == "title":
                    placeholder \
                        = slide.placeholders[PPTX_PLACEHOLDER_TITLE_TITLE]
                elif fields[0].strip() == "subtitle":
                    placeholder \
                        = slide.placeholders[PPTX_PLACEHOLDER_TITLE_SUBTITLE]
                else:
                    invalid = 1
            elif slide_type == "section":
                if fields[0].strip() == "title":
                    placeholder \
                        = slide.placeholders[PPTX_PLACEHOLDER_SECTION_TITLE]
                elif fields[0].strip() == "subtitle":
                    placeholder \
                        = slide.placeholders[PPTX_PLACEHOLDER_SECTION_SUBTITLE]
                else:
                    invalid = 1
            elif slide_type == "1content":
                if fields[0].strip() == "title":
                    placeholder \
                        = slide.placeholders[PPTX_PLACEHOLDER_1CONTENT_TITLE]
                elif fields[0].strip() == "content0":
                    placeholder \
                        = slide.placeholders[PPTX_PLACEHOLDER_1CONTENT_CONTENT0]
                else:
                    invalid = 1
            elif slide_type == "2content":
                if fields[0].strip() == "title":
                    placeholder \
                        = slide.placeholders[PPTX_PLACEHOLDER_2CONTENT_TITLE]
                elif fields[0].strip() == "content0":
                    placeholder \
                        = slide.placeholders[PPTX_PLACEHOLDER_2CONTENT_CONTENT0]
                elif fields[0].strip() == "content1":
                    placeholder \
                        = slide.placeholders[PPTX_PLACEHOLDER_2CONTENT_CONTENT1]
                else:
                    invalid = 1
            else:
                invalid = 1

            if invalid == 1:
                print("Error: Invalid field 0 \"" + fields[0]
                        + "\" for slide layout \"" + slide_type
                        + "\" (line {})\n".format(line_idx))
                raise

            if len(fields) < 2:
                print("Error: content line of slide layout \"" + slide_type
                        + "\" must contain more than 1 field. "
                        + "(line {})\n".format(line_idx))
                raise

            # Determine field 1 for a placeholder line
            if fields[1].strip() == "string":
                if len(fields) < 3:
                    print("Error: content line of slide layout \"" + slide_type
                            + "\" must contain more than 1 field. "
                            + "(line {})\n".format(line_idx))
                    raise

                string = re.sub(r'\"', '', fields[2])
                placeholder.text = string

            elif fields[1].strip() == "image":
                if len(fields) < 3:
                    print("Error: content line of slide layout \"" + slide_type
                            + "\" must contain more than 1 field. "
                            + "(line {})\n".format(line_idx))
                    raise

                string = re.sub(r'\"', '', fields[2])
                string = image_path + "/" + string
                if pathlib.Path(string).exists():
                    if pathlib.Path(string) in images:
                        images.remove(pathlib.Path(string))
                    insert_image(string, placeholder, slide)
                else:
                    invalid_images.append(string)
                    placeholder.text = "Image Not Found: " + string

            elif fields[1].strip() == "date":
                placeholder.text = datetime.datetime.now().strftime("%B %d, %Y")

            else:
                invalid = 1

            if invalid == 1:
                print("Error: Invalid field 1 \"" + fields[0]
                        + "\" for slide layout \"" + slide_type
                        + "\" (line {})\n".format(line_idx))
                raise

        # Invalid operation
        else:
            print("Error: Invalid field 0 \"" + fields[0] + 
                    "\" (line {})\n".format(line_idx))
            raise

    # Print invalid image paths
    if len(invalid_images) > 0:
        print("Invalid Image Paths: ")
        for path in invalid_images:
            print("  " + path)

    # Print unused image paths
    if len(images) > 0:
        print("Unused Image Paths: ")
        for path in images:
            print("  " + str(path))

# Function: insert_image()
# Description: Insert image into slide at position of placeholder
# while preserving the aspect ratio of the image. This function
# replaces the placeholder with the image, deleting the placeholer.
#
# Based on replace_with_image() by sclem
#   from https://github.com/scanny/python-pptx/issues/176
#
def insert_image(image, placeholder, slide):
    pic = slide.shapes.add_picture(image, placeholder.left, placeholder.top)

    # calculate max width/height for target size
    ratio = min(placeholder.width  / float(pic.width), placeholder.height / float(pic.height))

    pic.height = int(pic.height * ratio)
    pic.width  = int(pic.width  * ratio)

    pic.left = placeholder.left + ((placeholder.width  - pic.width)/2)
    pic.top  = placeholder.top  + ((placeholder.height - pic.height)/2)

    elem = placeholder.element
    elem.getparent().remove(elem)

    return pic

# Run program
main()
