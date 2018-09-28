#!/usr/bin/env python

# Force python XML parser not faster C accelerators
# because we can't hook the C implementation
import sys
sys.modules['_elementtree'] = None

import xml.etree.ElementTree as ET
import datetime
import os
import re
import argparse
from pptx import Presentation
from pprint import pprint


try:
    import pathlib
except ImportError:
    import pathlib2 as pathlib

def main():

    # Create presentation preprocessor and parse input
    ppp = PresentationPreprocessor("test/example.xml")


class PresentationPreprocessor:
    """This is the presentation preprocessor class.

    This class is used to preprocess and store information from an input
    XML file for the pptx-creator project.

    """

    def __init__(self, source=None):
        if source:
            self.parse(source)

    def parse(self, source):
        """ Load external XML document into the preprocessor.

        Load XML document and parse for pptx presentation format. A tree is
        created to store all of the presentation information.

        Args:
            source: The source XML document

        """

        # Open xml input file
        self.source = source
        etree = ET.parse(source, parser=LineNumberingParser())

        # Initialize parsing structures
        self.tree = PreprocessorEntry("_root_")
        self.var_stack = VariableStack()

        # Create tree, starting at the root
        self._process_element(etree.getroot(), self.tree)
        print(self.tree)

    def _process_element(self, elem, parent_entry, parent_elem=None):
        """ Process element and sub-elements.

        This function creates new PreprocessorEntry object for the specified
        element from the ElementTree, performs preprocessing tasks for
        the element, recursively calls this function for each of the element's
        attributes, recursively calls this function for each sub-element, and
        performs any postprocessing tasks for the element.

        Args:
            elem: ElementTree element to be processed
            parent_entry: Parent entry that is to hold the new PreprocessorEntry
                being created. The new PreprocessorEntry will be added to the
                end of the parent's data array.

        Kwargs:
            parent_elem: This is set when the current element is a temporary
                element created to represent an attribute of the parent element
                in the XML file.

        """


        # Create entry for element
        if not parent_elem is None:
            elem_entry = PreprocessorEntry(elem.tag,
                    parent=parent_entry, elem=parent_elem, is_attrib=True)
        else:
            elem_entry = PreprocessorEntry(elem.tag,
                    parent=parent_entry, elem=elem)

        # Perform tasks on element before calling children
        self._preprocess_element(elem_entry)

        # Process each attribute as a sub-element
        for item in elem.items():
            attrib = ET.Element(item[0])
            attrib.text = item[1]
            self._process_element(attrib, elem_entry, parent_elem=elem)

        # Add text as _value data under element
        elem_entry.add_text(elem.text)

        # Process subelements
        for child in elem:
            # Process child element
            self._process_element(child, elem_entry)

            # Get text after sub element
            elem_entry.add_text(child.tail)

        # Perform tasks on element after creating children
        self._postprocess_element(elem_entry)

    def _preprocess_element(self, elem_entry):
        """ Perform preprocessing for element.

        This is called by the _process_element function for each
        element before its sub-elements are processed.

        The following operations are performed by this function:
            - push new scope onto the variable stack

        Args:
            elem_entry: element entry in tree

        """

        # Append new dict onto the var_stack for the current scope variables
        self.var_stack.push()


    def _postprocess_element(self, elem_entry):
        """ Perform postprocessing for element.

        This is called by the _process_element function for each
        element after its sub-elements are processed.

        The following operations are performed by this function:
            - pop the current scope from the variable stack
            - process any get, mod, set elements
            - process any prepend, append elements (attributes)

        Args:
            elem_entry: element entry in tree

        """

        # Initialize values
        get_mod_set = None
        append_prepend = None

        # Pop current dict to return to scope of parent element
        self.var_stack.pop()

        # Determine variable action
        if (elem_entry.tag == "get"):
            get_mod_set = elem_entry.tag
        elif (elem_entry.tag == "mod"):
            get_mod_set = elem_entry.tag
        elif (elem_entry.tag == "set"):
            get_mod_set = elem_entry.tag
        elif (elem_entry.tag == "append"):
            append_prepend = elem_entry.tag
        elif (elem_entry.tag == "prepend"):
            append_prepend = elem_entry.tag


        # Get variable name being retrieved or modified
        if append_prepend:
            # Get variable name and lookup value
            var = elem_entry.get_values(join=True)
            try:
                val = self.var_stack.get(var)
            except ValueError as err:
                raise ValueError("var name \"{}\" does not exist; must set "\
                        "var before using \"{}\" element\n{}".format(var,
                            append_prepend, self.error_info(elem_entry))) \
                        from None

            # Create new PreprocessorValue object and add parent
            pp_val = PreprocessorValue(value=val)
            pp_val.parent = elem_entry.parent

            # Append or prepend data to parent's data array
            if append_prepend == "append":
                pp_val.parent.data.append(pp_val)
            elif append_prepend == "prepend":
                pp_val.parent.data.insert(0, pp_val)
            else:
                # Should never get here
                assert False

            # Delete entry from tree since we're evaluating it now
            elem_entry.delete()

        elif get_mod_set:

            # Find variable name
            var_array = elem_entry.get_values(tag="var",join=True)

            if (len(var_array) == 0):
                raise ValueError("cannot find name of var in \"{}\" "\
                        "element\n{}".format(get_mod_set, 
                            self.error_info(elem_entry)))

            if (len(var_array) > 1):
                raise ValueError("name of var specified {} times in \"{}\" "\
                        "element; can only provide one var name\n{}"\
                        "".format(len(var_array), get_mod_set,
                            self.error_info(elem_entry)))

            # Get variable and value
            var = var_array[0]
            val = elem_entry.get_values(join=True)

            # Ensure variable name is not empty string
            if (var == ""):
                raise ValueError("var name cannot be empty string in {} "\
                        "element\n{}".format(get_mod_set,
                            self.error_info(elem_entry)))

            # Get, mod, set value
            if (get_mod_set == "get"):
                val = self.var_stack.get(var)
                elem_entry.to_value(val)
            elif (get_mod_set == "mod"):
                val = self.var_stack.mod(var, val)
                elem_entry.delete()
            elif (get_mod_set == "set"):
                val = self.var_stack.set(var, val)
                elem_entry.delete()

    def error_info(self, elem_entry):
        return "\tFile \"{}\" {}".format(self.source, elem_entry.error_info())


class VariableStack:
    """ This is the variable stack class.

    This class is used to create and manage the variable stack of the input
    XML file. The variables are created and accessed in the XML using the
    set, mod, and get elements and the prepend and append attributes.

    """

    def __init__(self):
        self.var_stack = []

    def push(self):
        self.var_stack.append({})

    def pop(self):
        self.var_stack.pop()

    def set(self, var, val):
        self.var_stack[-1][var] = val

    def mod(self, var, val):
        self.find_dict(var)[var] = val

    def get(self, var):
        return self.find_dict(var)[var]

    def find_dict(self, var):
        for var_dict in self.var_stack:
            if var in var_dict:
                return var_dict

        raise ValueError("var \"{}\" not found in variable stack".format(var))


class PreprocessorEntry:
    """ This is the preprocess entry class.

    The preprocess entry object is used by the PresentationPreprocessor
    class to hold elements and attributes in a standard data structure.

    An entry has a tag attribute and an entries attribute.
        - tag:  string containing attribute or element name from XML file
        - data: array of additional entries or values

    An object can be constructed with or without a value, as shown below
        entry = PreprocessorEntry("my_tag")
            Returns:
                entry.tag  == "my_tag"
                entry.data == []

        entry = PreprocessorEntry("my_tag",value="my_value")
            Returns:
                entry.tag  == "my_tag"
                entry.data[0].value == "my_value"

    """
    # Used to differentiate PreprocessorEntry and PreprocessorValue objects
    which = "entry"

    def __init__(self, tag, parent=None, value=None, elem=None, is_attrib=False):
        """ Initialize new PreprocessorEntry object.

        Create object with tag and empty data array. The capability is also
        provided to add a value associated with the enty, by setting value.
        When value is set, a new PreprocessorValue object is created with
        the value and added as the first entry of the ProprocessorEntry
        data array.

        The caller also has the option to associate an ElementTree element
        with this entry using the elem keyword argument. This can be done
        to provide clearer error messages and to aid in debugging. The
        is_attrib keyword argument can be used to indicate whether this entry
        is the element itself or an attribute of the element.

        Args:
            tag: The name of the tag for the entry which corresponds to the
                element tag or the attribute name.

        Kwargs:
            parent: The parent entry that is to hold the newly created entry.
                The new entry will be appended to the end of the data array
                of the parent entry.
            value: The value associated with the entry, typically this is
                used when adding an attribute entry so that both the name
                and value of the attribute can be added at once.
            elem: The ElementTree element corresponding to this entry. This
                is used for printing helpful error messages.
            is_attrib: True/False to indicate whether the entry is an
                attribute or an element in the original XML file.

        """

        # Initialize data
        self.tag = tag
        self.data = []
        self.elem = elem
        self.is_attrib = is_attrib

        # When value is specified create PreprocessorValue in data array
        if value:
            self.add_value(value)

        # Add entry to end of data array for parent
        if parent:
            self.parent = parent
            parent.data.append(self)

    def __repr__(self):
        return "<{} '{}' at {}>"\
                "".format( self.__class__.__name__, self.tag, hex(id(self)))

    def __str__(self):
        ret  = "<{}: \'{}\'".format(self.which, self.tag)

        if self.elem is None:
            ret += ">"
        else:
            ret += ", line: {}>".format(self.elem._start_line_number)

        for pp in self.data:
            for line in str(pp).splitlines():
                ret += "\n   {}".format(line)

        return ret

    def add_value(self, value):
        """ Add value to entry.

        Add a new PreprocessorValue entry to the current PreprocessorEntry
        object's data array.

        Args:
            value: value to be added to entry data array

        """
        PreprocessorValue(self, value=value)

    def add_text(self, text):
        """ Add text to entry.

        Perform text processing (e.g. remove leading and trailing spaces
        on string.  Then add a new PreprocessorValue entry to the current
        PreprocessorEntry object's data array with the value of the
        processed text.

        Args:
            text: The text string to be added to the tree

        """

        # Don't add when text is None
        if (text is None):
            return

        # TODO may need to make this more intelligent
        # Remove all leading and trailing whitespace characters
        text = text.strip()

        # Only add text entry when text is not empty string
        if (text != ""):
            self.add_value(text)

    def get_values(self, tag=None, join=False):
        """ Return value(s) of entry

        Get all values (PreprocessorValue object data) in the current
        entry's data array. When a tag is specified, all value entries
        associated with the tag are returned, as an array. When the
        join argument is True, all the values associated with
        a specific entry are joined together. Therefore, when
        tag is unspecified and join is True, all the values of the current
        entry are joined and returned as a single value. When join
        is False, an array of values is returned. When tag is specified and
        join is True, an array of the joined values for each entry
        associated with the tag are returned. When the tag is specified
        and join is False, an array of the arrays of values for each entry
        associated with the tag are returned.

        Kwargs:
            tag: name of tag to retreive values from
            join: indicates whether to join values

        Return:
            if tag=None and join=True:  joined values
            if tag=None and join=False: array of values
            if tag specified and join=True:  array of joined values
            if tag specified and join=False: array of arrays of values

        """

        # Get values from entries matching tag in data array
        if tag:
            return [pp.get_values(join=join) for pp in self.data
                    if pp.which == "entry" and pp.tag == tag]

        # Get values in data array
        values = [pp.value for pp in self.data if pp.which == "value"]

        if join:
            return ''.join(values)
        else:
            return values

    def delete(self):
        """ Delete entry from parent's data array.

        Effectively deletes the element from the whole entry tree by removing
        the reference from the parent's data array.
        """

        self.parent.data.remove(self)

    def to_value(self, val):
        """ Convert the current entry to a value.

        Delete the current PreprocessorEntry from the parent's data array
        and replace it with a ProprocessorValue with the value specified
        by val.
        """

        # Get current index and replace it with new PreprocessorValue object
        idx = self.parent.data.index(self)
        pp_val = PreprocessorValue(value=val)
        self.parent.data[idx] = pp_val

        # Update parent pointer in new pp_val object
        pp_val.parent = self.parent

    def error_info(self):
        return "in element \"{}\", line {}"\
                "".format(self.parent.tag, self.elem._start_line_number)

class PreprocessorValue:
    """ This is the preprocess value class.

    The preprocess value object is used by the PresentationPreprocessor
    class to hold values of elements and data for an entry.

    A value object has a value attribute, which holds the text value of
    elements and attributes from the XML file.

    """
    # Used to differentiate PreprocessorEntry and PreprocessorValue objects
    which = "value"

    def __init__(self, parent=None, value=None):
        """ Initialize new PreprocessorValue object.

        Create object with parent and value. The value is stored in the
        value variable for the object. The parent specifies the parent
        PreprocessorEntry where the new new PreprocessorValue should be
        added. This function will append the new PreprocessorValue to the
        end of the parent's data array.
        When value is set, a new PreprocessorValue object is created with

        Kwargs:
            parent: The parent entry that is to hold the newly created value.
                The new value will be appended to the end of the data array
                of the parent entry.
            value: The value of the PreprocessorValue object.

        """

        if value:
            self.value = value

        # Add entry to end of data array for parent
        if parent:
            self.parent = parent
            parent.data.append(self)

    def __repr__(self):
        return "<{} '{}' at {}>"\
                "".format( self.__class__.__name__, self.value, hex(id(self)))

    def __str__(self):
        ret  = "<{}: \"{}\">".format(self.which, self.value)
        return ret



# From Duncan Harris on stack overflow: https://stackoverflow.com/questions/
#   6949395/is-there-a-way-to-get-a-line-number-from-an-elementtree-element
class LineNumberingParser(ET.XMLParser):
    def _start(self, *args, **kwargs):
        # Here we assume the default XML parser which is expat
        # and copy its element position attributes into output Elements
        element = super(self.__class__, self)._start(*args, **kwargs)
        element._start_line_number = self.parser.CurrentLineNumber
        element._start_column_number = self.parser.CurrentColumnNumber
        element._start_byte_index = self.parser.CurrentByteIndex
        return element

    def _end(self, *args, **kwargs):
        element = super(self.__class__, self)._end(*args, **kwargs)
        element._end_line_number = self.parser.CurrentLineNumber
        element._end_column_number = self.parser.CurrentColumnNumber
        element._end_byte_index = self.parser.CurrentByteIndex
        return element



# Run program
main()















