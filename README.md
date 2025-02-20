# LoPolyfill

Copyright (C) 2025 Julien Férard <https://github.com/jferard>

A polyfill for LibreOffice Calc 7.2.

Under GPL v3.

**The documentation of the provided functions and their parameters is taken
from the [LibreOffice help pages](https://help.libreoffice.org) (
[Mozilla Public License v2.0](https://www.mozilla.org/MPL/))**

## Summary

LoPolyfill is an extension that aims provide recent Calc functions for
LibreOffice 7.2 (Python 3.8) and following. It is intended for
organisations that have a large computer fleet and don't want to update
LibreOffice for various reasons.

**This extension does not try to provide MS Excel functions (see
[Lox365](https://github.com/goosepirate/lox365) for MS Excel functions).**

This extension favor compatibility over speed: the behavior of the functions
should be very close to the actual LibreOffice functions.

## Important

## Usage

Once the extension is installed, the functions are available in the **Add-In**
category, and have the `LOP.` (**L**ibre**O**ffice**P**olyfill) prefix.

Those functions are available in LibreOffice Calc since the version 24.8
(Aug. 2024) and, soon, 25.8 (Aug. 2025).

## Calc Functions

| Function       | LibreOffice                                                                                                                                                | Excel                                                                                                             | Comment                                                |
|----------------|------------------------------------------------------------------------------------------------------------------------------------------------------------|-------------------------------------------------------------------------------------------------------------------|--------------------------------------------------------|
| LOP.FILTER     | [FILTER](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_filter.html) (*)                                                                     | [FILTER](https://support.microsoft.com/en-us/office/filter-function-f4f7cb66-82eb-4767-8f7c-4877ad80c759)         |                                                        |                                                        |
| -              | [LET](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_let.html) ([24.8](https://wiki.documentfoundation.org/ReleaseNotes/24.8#New_functions)) | [LET](https://support.microsoft.com/en-us/office/let-function-34842dd8-b92b-4d3f-b325-b8b8f9908999)               | Won't be implemented                                   |
| LOP.RANDARRAY  | [RANDARRAY](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_randarray.html) (*)                                                               | [RANDARRAY](https://support.microsoft.com/en-us/office/randarray-function-21261e55-3bec-4885-86a6-8b0a47fd4d33)   |                                                        |
| LOP.SEQUENCE   | [SEQUENCE](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_sequence.html) (*)                                                                 | [SEQUENCE](https://support.microsoft.com/en-us/office/sequence-function-57467a98-57e0-4817-9f14-2eb78519ca90)     |                                                        |
| LOP.SORT       | [SORT](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_sort.html) (*)                                                                         | [SORT](https://support.microsoft.com/en-us/office/sort-function-22f63bd0-ccc8-492f-953d-c20e8e44b86c)             | Uses document collator                                 |
| LOP.SORTBY     | [SORTBY](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_sortby.html) (*)                                                                     | [SORTBY](https://support.microsoft.com/en-us/office/sortby-function-cd2d7a62-1b93-435c-b561-d6a35134f28f)         | Uses document collator                                 |
| LOP.UNIQUE     | [UNIQUE](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_unique.html) (*)                                                                     | [UNIQUE](https://support.microsoft.com/en-us/office/unique-function-c5ab87fd-30a3-4ce9-9d1a-40204fb85e1e)         |                                                        |
| LOP.XLOOKUP    | [XLOOKUP](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_xlookup.html) (*)                                                                   | [XLOOKUP](https://support.microsoft.com/en-us/office/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929)       | Uses document collator, wildcard/regex not implemented |
| LOP.XMATCH     | [XMATCH](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_xmatch.html) (*)                                                                     | [XMATCH](https://support.microsoft.com/en-us/office/xmatch-function-d966da31-7a6b-4a13-a1c6-5a33ed6a0312)         | Uses document collator, wildcard/regex not implemented |
| LOP.CHOOSECOLS | [CHOOSECOLS](https://help.libreoffice.org/25.8/en-US/text/scalc/01/func_choosecols.html) (**)                                                              | [CHOOSECOLS](https://support.microsoft.com/en-us/office/choosecols-function-bf117976-2722-4466-9b9a-1c01ed9aebff) |                                                        |
| LOP.CHOOSEROWS | [CHOOSEROWS](https://help.libreoffice.org/25.8/en-US/text/scalc/01/func_chooserows.html) (**)                                                              | [CHOOSEROWS](https://support.microsoft.com/en-us/office/chooserows-function-51ace882-9bab-4a44-9625-7274ef7507a3) |                                                        |
| LOP.DROP       | [DROP](https://help.libreoffice.org/25.8/en-US/text/scalc/01/func_drop.html) (**)                                                                          | [DROP](https://support.microsoft.com/en-us/office/drop-function-1cb4e151-9e17-4838-abe5-9ba48d8c6a34)             |                                                        |
| LOP.TAKE       | [TAKE](https://help.libreoffice.org/25.8/en-US/text/scalc/01/func_take.html) (**)                                                                          | [TAKE](https://support.microsoft.com/en-us/office/take-function-25382ff1-5da1-4f78-ab43-f33bd2e4e003)             |                                                        |
| LOP.EXPAND     | [EXPAND](https://help.libreoffice.org/25.8/en-US/text/scalc/01/func_expand.html) (**)                                                                      | [EXPAND](https://support.microsoft.com/en-us/office/expand-function-7433fba5-4ad1-41da-a904-d5d95808bc38)         |                                                        |
| LOP.HSTACK     | [HSTACK](https://help.libreoffice.org/25.8/en-US/text/scalc/01/func_hstack.html) (**)                                                                      | [HSTACK](https://support.microsoft.com/en-us/office/hstack-function-98c4ab76-10fe-4b4f-8d5f-af1c125fe8c2)         |                                                        |
| LOP.VSTACK     | [VSTACK](https://help.libreoffice.org/25.8/en-US/text/scalc/01/func_vstack.html) (**)                                                                      | [VSTACK](https://support.microsoft.com/en-us/office/vstack-function-a4b86897-be0f-48fc-adca-fcc10d795a9c)         |                                                        |
| LOP.TOCOL      | [TOCOL](https://help.libreoffice.org/25.8/en-US/text/scalc/01/func_tocol.html) (**)                                                                        | [TOCOL](https://support.microsoft.com/en-us/office/tocol-function-22839d9b-0b55-4fc1-b4e6-2761f8f122ed)           |                                                        |
| LOP.TOROW      | [TOROW](https://help.libreoffice.org/25.8/en-US/text/scalc/01/func_torow.html) (**)                                                                        | [TOROW](https://support.microsoft.com/en-us/office/torow-function-b90d0964-a7d9-44b7-816b-ffa5c2fe2289)           |                                                        |
| LOP.WRAPCOLS   | [WRAPCOLS](https://help.libreoffice.org/25.8/en-US/text/scalc/01/func_wrapcols.html) (**)                                                                  | [WRAPCOLS](https://support.microsoft.com/en-us/office/wrapcols-function-d038b05a-57b7-4ee0-be94-ded0792511e2)     |                                                        |
| LOP.WRAPROWS   | [WRAPROWS](https://help.libreoffice.org/25.8/en-US/text/scalc/01/func_wraprows.html) (**)                                                                  | [WRAPROWS](https://support.microsoft.com/en-us/office/wraprows-function-796825f3-975a-4cee-9c84-1bbddf60ade0)     |                                                        |

(*) [LibreOffice 24.8](https://wiki.documentfoundation.org/ReleaseNotes/24.8#New_functions)

(**) [LibreOffice 25.8](https://wiki.documentfoundation.org/ReleaseNotes/25.8#New_functions)