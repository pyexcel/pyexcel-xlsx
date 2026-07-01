"""Test package initialization and compatibility patches."""

from xml.etree import ElementTree

# Compatibility for older openpyxl versions on modern Python.
# Python 3.9+ removed ElementTree.getiterator; openpyxl<2.5 may still call it.
if not hasattr(ElementTree.ElementTree, "getiterator"):
    ElementTree.ElementTree.getiterator = ElementTree.ElementTree.iter
