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

| Function                                                                                                                                                                                                | LibreOffice version                                                         | Comment                                                |
|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|-----------------------------------------------------------------------------|--------------------------------------------------------|
| [FILTER](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_filter.html) ([MS Excel](https://support.microsoft.com/en-us/office/filter-function-f4f7cb66-82eb-4767-8f7c-4877ad80c759))        | [24.8](https://wiki.documentfoundation.org/ReleaseNotes/24.8#New_functions) |                                                        |
| [LET](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_let.html) ([MS Excel](https://support.microsoft.com/en-us/office/let-function-34842dd8-b92b-4d3f-b325-b8b8f9908999))                 | [24.8](https://wiki.documentfoundation.org/ReleaseNotes/24.8#New_functions) | Won't be implemented                                   |
| [RANDARRAY](https://help.libreoffice.org/25.2/en-US/text/scalc/01/func_randarray.html) ([MS Excel](https://support.microsoft.com/en-us/office/randarray-function-21261e55-3bec-4885-86a6-8b0a47fd4d33)) | [24.8](https://wiki.documentfoundation.org/ReleaseNotes/24.8#New_functions) |                                                        |
| [SEQUENCE](https://help.libreoffice.org/25.2/en-US/text/scalc/01/func_sequence.html) ([MS Excel](https://support.microsoft.com/en-us/office/sequence-function-57467a98-57e0-4817-9f14-2eb78519ca90))    | [24.8](https://wiki.documentfoundation.org/ReleaseNotes/24.8#New_functions) |                                                        |
| [SORT](https://help.libreoffice.org/25.2/en-US/text/scalc/01/func_sort.html) ([MS Excel](https://support.microsoft.com/en-us/office/sort-function-22f63bd0-ccc8-492f-953d-c20e8e44b86c))                | [24.8](https://wiki.documentfoundation.org/ReleaseNotes/24.8#New_functions) | Uses document collator                                 |
| [SORTBY](https://help.libreoffice.org/25.2/en-US/text/scalc/01/func_sortby.html) ([MS Excel](https://support.microsoft.com/en-us/office/sortby-function-cd2d7a62-1b93-435c-b561-d6a35134f28f))          | [24.8](https://wiki.documentfoundation.org/ReleaseNotes/24.8#New_functions) | Uses document collator                                 |
| [UNIQUE](https://help.libreoffice.org/25.2/en-US/text/scalc/01/func_unique.html) ([MS Excel](https://support.microsoft.com/en-us/office/unique-function-c5ab87fd-30a3-4ce9-9d1a-40204fb85e1e))          | [24.8](https://wiki.documentfoundation.org/ReleaseNotes/24.8#New_functions) | TO FIX: does not use document collator                 |
| [XLOOKUP](https://help.libreoffice.org/25.2/en-US/text/scalc/01/func_xlookup.html) ([MS Excel](https://support.microsoft.com/en-us/office/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929))       | [24.8](https://wiki.documentfoundation.org/ReleaseNotes/24.8#New_functions) | Uses document collator, wildcard/regex not implemented |
| [XMATCH](https://help.libreoffice.org/25.2/en-US/text/scalc/01/func_xmatch.html) ([MS Excel](https://support.microsoft.com/en-us/office/xmatch-function-d966da31-7a6b-4a13-a1c6-5a33ed6a0312))          | [24.8](https://wiki.documentfoundation.org/ReleaseNotes/24.8#New_functions) | Uses document collator, wildcard/regex not implemented |
| [CHOOSECOLS]() ([MS Excel](https://support.microsoft.com/en-us/office/choosecols-function-bf117976-2722-4466-9b9a-1c01ed9aebff))                                                                        | [25.8](https://wiki.documentfoundation.org/ReleaseNotes/25.8#New_functions) | Not implemented                                        |
| [CHOOSEROWS]() ([MS Excel](https://support.microsoft.com/en-us/office/chooserows-function-51ace882-9bab-4a44-9625-7274ef7507a3))                                                                        | [25.8](https://wiki.documentfoundation.org/ReleaseNotes/25.8#New_functions) | Not implemented                                        |
| [DROP]() ([MS Excel](https://support.microsoft.com/en-us/office/drop-function-1cb4e151-9e17-4838-abe5-9ba48d8c6a34))                                                                                    | [25.8](https://wiki.documentfoundation.org/ReleaseNotes/25.8#New_functions) | Not implemented                                        |
| [EXPAND]() ([MS Excel](https://support.microsoft.com/en-us/office/expand-function-7433fba5-4ad1-41da-a904-d5d95808bc38))                                                                                | [25.8](https://wiki.documentfoundation.org/ReleaseNotes/25.8#New_functions) | Not implemented                                        |
| [HSTACK]() ([MS Excel](https://support.microsoft.com/en-us/office/hstack-function-98c4ab76-10fe-4b4f-8d5f-af1c125fe8c2))                                                                                | [25.8](https://wiki.documentfoundation.org/ReleaseNotes/25.8#New_functions) | Not implemented                                        |
| [TAKE]() ([MS Excel](https://support.microsoft.com/en-us/office/take-function-25382ff1-5da1-4f78-ab43-f33bd2e4e003))                                                                                    | [25.8](https://wiki.documentfoundation.org/ReleaseNotes/25.8#New_functions) | Not implemented                                        |
| [TOCOL]() ([MS Excel](https://support.microsoft.com/en-us/office/tocol-function-22839d9b-0b55-4fc1-b4e6-2761f8f122ed))                                                                                  | [25.8](https://wiki.documentfoundation.org/ReleaseNotes/25.8#New_functions) | Not implemented                                        |
| [TOROW]() ([MS Excel](https://support.microsoft.com/en-us/office/torow-function-b90d0964-a7d9-44b7-816b-ffa5c2fe2289))                                                                                  | [25.8](https://wiki.documentfoundation.org/ReleaseNotes/25.8#New_functions) | Not implemented                                        |
| [VSTACK]() ([MS Excel](https://support.microsoft.com/en-us/office/vstack-function-a4b86897-be0f-48fc-adca-fcc10d795a9c))                                                                                | [25.8](https://wiki.documentfoundation.org/ReleaseNotes/25.8#New_functions) | Not implemented                                        |
| [WRAPCOLS]() ([MS Excel](https://support.microsoft.com/en-us/office/wrapcols-function-d038b05a-57b7-4ee0-be94-ded0792511e2))                                                                            | [25.8](https://wiki.documentfoundation.org/ReleaseNotes/25.8#New_functions) | Not implemented                                        |
| [WRAPROWS]() ([MS Excel](https://support.microsoft.com/en-us/office/wraprows-function-796825f3-975a-4cee-9c84-1bbddf60ade0))                                                                            | [25.8](https://wiki.documentfoundation.org/ReleaseNotes/25.8#New_functions) | Not implemented                                        |
