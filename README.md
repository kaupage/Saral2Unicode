# Saral2Unicode

This script coverts Saral Font (Devanagari Typing Font) to Unicode (e.g. used in Google input tools)

## Background -
Saral font was created back in the days when Devanagari (देवनागरी) characters were not added in Unicode.
Unicode block was added in Unicode in 1999.
https://en.wikipedia.org/wiki/Devanagari_(Unicode_block)

Pre-1999, Saral font was widely used for typing. Saral mapped English block in Unicode to Devanagari font.
(e.g. A = अ and a = ा)

## Font Mapping
First step in font conversion is to create mapping file between two fonts.
'Saral - Unicode Mapping' Excel file contains mapping for various font groups.
It covers diacritical and vowel signs in Devanagari.

## Script Info
  mergeArrays and IsInArray are simple array processing functions
  
* SaralToUnicode() is the conversion script
* Mapping array are written next to each other with font name prefix (e.g. level0_Uni maps to level0_Saral)
* line 146 detects if it is a ि , it needs to be treated differently as it is positioned post letter in Saral and pre letter in Unicode
* For numbers and symbols, the replacement happens post character replacement.
