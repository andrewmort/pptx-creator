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
import xml.etree.ElementTree as ET
import datetime
import os
import re
import argparse
import sys

try:
    import pathlib
except ImportError:
    import pathlib2 as pathlib


def main():
    # Parse arguments to get paths
    path_input, path_output, path_xml, path_pptx = parse_arguments()

    # Interpret template xml file
    template = get_template(path_xml)

    # Parse input lines to create presentation
    create_presentation(path_input, path_output, path_pptx, template)


# Parse input arguments
def parse_arguments():
    global verbose

    # Setup argument parser
    parser = argparse.ArgumentParser(description="" \
        "This program creates a powerpoint (pptx) file based on an xml"      \
        " definition file and a template. A template consists of a template" \
        " pptx file and a template xml file, both of which typically reside" \
        " in a template directory.")
    parser.add_argument("input",  help="xml definition file")
    parser.add_argument("-o", "--output", help="pptx output file")
    parser.add_argument("-t", "--template",
        help="template directory, containing pptx and xml files, "\
        "file names should either be template.pptx and template.xml or "\
        "<dirname>.pptx and <dirname>.xml"\
        "(specify individual file locations  with --pptx and --xml)")
    parser.add_argument("-p", "--pptx",
        help="template pptx file (overrides location of pptx from --template)")
    parser.add_argument("-x", "--xml",
        help="template xml file (overrides location of xml from --template)")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args()

    # Get arguments
    verbose           = args.verbose
    filename_input    = args.input
    filename_output   = args.output
    filename_template = args.template
    filename_pptx     = args.pptx
    filename_xml      = args.xml

    # Get pathlib object for input
    path_input = pathlib.Path(filename_input)

    if verbose:
        print("INFO: Input path " + str(path_input) + ".")

    # Get pathlib object for output
    if (filename_output):
        path_output = pathlib.Path(filename_output)
    else:
        # Path to output file with same name as input file in pwd
        path_output = path_input.with_suffix('.pptx').name

    if verbose:
        print("INFO: Output path " + str(path_output) + ".")

    # Get pathlib object for template
    if (filename_template):
        path_template = pathlib.Path(filename_template)
        template_names = (path_template.name, 'template')

    # Get xml path
    if (filename_xml):
        path_xml = pathlib.Path(filename_xml)
    else:
        if (not path_template):
            raise ValueError("No xml or template path specified!")

        # Search for valid xml paths
        for name in template_names:
            path_temp = path_template.joinpath(name + '.xml')
            if path_temp.exists():
                path_xml = path_temp
                break

        # Unable to find valid xml file
        if (not path_xml):
            raise ValueError("No valid template xml file found!")

    if verbose:
        print("INFO: XML template path " + str(path_xml) + ".")

    # Get pptx path
    if (filename_pptx):
        path_pptx = pathlib.Path(filename_pptx)
    else:
        if (not path_template):
            raise ValueError("No pptx or template path specified!")

        # Search for valid xml paths
        for name in template_names:
            path_temp = path_template.joinpath(name + '.pptx')
            if path_temp.exists():
                path_pptx = path_temp
                break

        # Unable to find valid xml file
        if (not path_pptx):
            raise ValueError("No valid template pptx file found!")

    if verbose:
        print("INFO: PPTX template path " + str(path_pptx) + ".")

    return path_input, path_output, path_xml, path_pptx

# Get template mapping from xml file
def get_template(path_xml):
    template = {}

    # Open xml template definition file
    xml_template = ET.parse(path_xml)

    # Create map from child (c) to parent (p)
    xml_parents = {c:p for p in xml_template.iter() for c in p}

    template_idx = -1
    layout_idx = -1
    ph_idx = -1

    # Iterate over xml elements to create template
    for xml_elem in xml_template.iter():
        if (xml_elem.tag == 'template'):
            template_idx += 1

            # Ensure template has no parent
            try:
                xml_parents[xml_elem]
            except KeyError:
                pass
            else:
                raise ValueError("Template should be top level element!")

            if (template_idx > 0):
                raise ValueError("Cannot have more than one template element!")

        elif (xml_elem.tag == 'layout'):
            # Update index values for error messages
            layout_idx += 1
            ph_idx = -1

            layout = xml_elem.get("name")
            index  = xml_elem.get("index")

            if (not layout):
                raise ValueError("Layout \"{}\" missing name attribute in xml "\
                    "template file {}"\
                    "".format(layout_idx, str(path_xml)))

            if (xml_parents[xml_elem].tag != 'template'):
                raise ValueError("Layout \"{}\" ({}) has parent node "\
                    "\"{}\" and should have parent node \"template\" in xml "\
                    " template file {}"\
                    "".format(layout_idx, layout, xml_parents[xml_elem].tag,
                        str(path_xml)))

            if (not index):
                raise ValueError("Layout \"{}\" ({}) missing index attribute "\
                    "in xml template file {}"\
                    "".format(layout_idx, layout, str(path_xml)))

            # Associate index with layout name
            template[layout] = {}
            template[layout]["idx"] = index

        elif (xml_elem.tag == 'placeholder'):
            # Update index values for error messages
            ph_idx += 1

            ph     = xml_elem.get("name")
            index  = xml_elem.get("index")

            if (not ph):
                raise ValueError("Placeholder \"{}\" in layout \"{}\" ({}) "\
                    "missing name attribute in xml template file {}"\
                    "".format(ph_idx, layout_idx, layout, str(path_xml)))

            if (xml_parents[xml_elem].tag != 'layout'):
                raise ValueError("Placeholder \"{}\" ({}) has parent node "\
                    "\"{}\" and should have parent node \"layout\" in xml "\
                    " template file {}"\
                    "".format(ph_idx, ph, xml_parents[xml_elem].tag,
                        str(path_xml)))

            if (not index):
                raise ValueError("Placeholder \"{}\" ({}) in layout \"{}\" "\
                    "({}) missing index attribute in xml template file {}"\
                    "".format(ph_idx, ph, layout_idx, layout, str(path_xml)))

            # Associate index with ph name for current layout
            template[layout]["ph"] = {}
            template[layout]["ph"][ph] = index

        else:
            raise ValueError("Invalid tag \"{}\" in xml template file {}"\
                "".format(xml_elem.tag, str(path_xml)))

    # Invalid template file if no layout tags are found
    if (layout_idx < 0):
        raise ValueError("No layout tags found in xml template file {}"\
            "".format(str(path_xml)))

    return template

def create_presentation(path_input, path_output, path_pptx, template):
    # Create presentation
    prs = Presentation(path_pptx)

    # Open xml input file
    xml_input = ET.parse(path_input)

    # Pre-process xml input file
    preprocess_input(prs, xml_input)

# This functions
#   - Create slides
#   - Create slide references
#   - Evaluate and replace variables
#   - Orgnanize data
def preprocess_input(prs, xml_input):
    var_stack = []
    elem_stack = []
    input_stack = []

    # Create map from child (c) to parent (p)
    xml_parents = {c:p for p in xml_input.iter() for c in p}

    level = -1

    # Process each child, creating tree of elements and attributes
    for child in xml_input.iter():
        # **** Determine action based on location in stack ****

        # First element in stack, add child
        if (len(elem_stack) == 0):
            append = 1

        # Parent is last element in stack, add child
        elif (xml_parents[child] == elem_stack[-1]):
            append = 1

        # Parent further up stack, remove children to parent, add child
        else:
            append = -elem_stack.index(xml_parents[child])

        preprocess_element(append, child, input_vars)

    # After loop, ensure all elements have been processed
    preprocess_element(-len(elem_stack), None, input_vars)


def preprocess_push(tag, pre_vars):
    tree_stack = pre_vars["tree_stack"]

    # Create new tree element
    new_elem = {}
    new_elem["tag"] = tag
    new_elem["data"] = []

    # Add new element as data of element on top of stack
    tree_stack[-1].append(new_elem)
    tree_stack.append(new_elem["data"])

def preprocess_pop(pre_vars)
    tree_stack = pre_vars["tree_stack"]
    var_stack  = pre_vars["var_stack"]

    # Remove element from top of stack
    old_elem = tree_stack.pop()

    # Operate on element based on tag
    if (old_elem["tag"] == "get")
        name = [item["data"] for item in old_elem["data"] if item["tag"] == "name"]

        if (len(name) != 1 or len(name[0]) != 1):
            #TODO improve error message
            raise ValueError("Must specify name with <get/> tag.")




    elif 

def get_varstac




#
# Function: preprocess_element
# Date: 9/18/2018
# Description: Process each presentation input xml element, creating a tree
#   of elements and attributes. To keep track of parents and heirarchy,
#   elements are added to a stack. The parameter append is used to operate
#   on the stack.
#
# Parameters:
#   append      when  1, element is child of previous element, add to stack
#               when <0, indicates elements to pop to get child's parent
#   child       current element to add to stack and process
#   input_vars  array of variables that are used to process and create the tree
#
def preprocess_element(append, child, input_vars):
    tree        = input_vars["tree"]
    path        = input_vars["path"]
    tree_stack  = input_vars["tree_stack"]
    var_stack   = input_vars["var_stack"]

    # Every element added as {elem_name, {data}}

    # Add each element and all it's attributes to tree
    if (append):
        # Create new element
        new_elem = []
        new_elem["tag"] = child.tag
        new_elem["data"] = {}

        # Add new element under previous
        tree_stack[-1].append(new_elem)
        tree_stack.append(new_elem["data"])

        # Get element's attributes
        for item in child.items():
            new_item = []
            new_item["tag"] = item[0]
            new_item["data"] = {item[1]}

            tree_stack[-1].append(new_elem)






    # Pop element off of stack
    else:

    # pop element
    #   - remove element from element stack
    #   - remove element location in tree from path
    #   - dereference variables for gets, append, prepend
    #   - set variable values for sets, mods
    #   - add slide and layout to list of slides
    #   - associate slide with tag/label

    # push element
    #   - add element to element stack
    #   - add eleemnt and all args to tree
    #   - add element to location in tree based on the path
    #   - update element path


    # Add element
    if (append > 0):
        elem_stack.append(child)
        var_stack.append({})
        level+=1

        # **** Process action ****

        #
        # If append
        #   - add element to stack
        #   - create new variable stack level
        #   - add element and attributes to tree
        # Else
        #   - pop element stack to parent
        #   - pop variable stack
        #   - process each element for variables
        #

        # Process actions
        if (append):
        else:
            for i in range(idx, len(elem_stack)-1):
                elem_stack.pop()
                var_stack.pop()
            elem_stack.append(child)
            var_stack.append({})
            level = idx + 1

        print ("  " * level + str(child))


def oldnewnew():
    # Initialize data strctures to hold input xml information
    prs_slides = []
    prs_slides_ref = {}
    slides_layout = []
    defines = {}

    slide_idx = -1

    # Create slides first so we can create links while adding content
    for xml_slide in xml_input.iter('slide'):
        # Update index values for error messages
        slide_idx += 1

        layout = xml_slide.get('layout')

        if (not layout):
            raise ValueError("Slide {} is missing layout attribute Layout {} missing name attribute in xml "\
                "template file {}"\
                "".format(layout_idx, str(path_xml)))

        layout_idx = template[layout]["idx"]


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

    # Save presetnation as filename_output
    prs.save(path_output)

# Use lines from input file to create presentation
def oldnew_make_presentation(prs, xml_tree):
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
        label = xml_slide.get('label')
        if (label):
            # Do not allow duplicate references
            if (label in slides_ref):
                print("Error: Label {} on slide {} (layout {}) already "
                        "exists".format(label,slide_idx+1,layout))
                raise

            prs_slides_ref[label] = slides[slide_idx]

        slide_idx += 1

    # Reset variable for loop
    slide_idx = 0

    # Loop through all xml elements in order
    for xml_child in xml_tree:
        1+1
        # Slide definitions
        #if (xml_child.tag == 'slide'):


        # Variable definitions
        #elif (xml_child.tag == 'define'):

        # Invalid elements
        #else:
            #print("Error: Invalid element \"" + xml_child.tag + "\"")
            #raise


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

class Preprocessor:
    """Preprocessor

    This class is used to preprocess and store information from an input
    XML file for the pptx-creator project.

    """

    def __init__(self):
        self.xml_input = None

    def parse(self, xml_path):
        """parse(xml_path)

        This function reads the xml file specified by xml_path and parses
        it into the presentation tree that can be accessed by the other
        class function
        """

        # Open xml input file
        self.xml_path = xml_path
        self.xml_input = ET.parse(xml_path)







