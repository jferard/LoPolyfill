import re
from typing import Any

# noinspection PyUnresolvedReferences
import uno
# noinspection PyUnresolvedReferences
from com.sun.star.beans import XPropertySet
# noinspection PyUnresolvedReferences
from com.sun.star.i18n import XCollator
# noinspection PyUnresolvedReferences
from com.sun.star.uno import XComponentContext


class MessageBoxType:
    # noinspection PyUnresolvedReferences
    from com.sun.star.awt.MessageBoxType import (MESSAGEBOX, )


class MessageBoxButtons:
    # noinspection PyUnresolvedReferences
    from com.sun.star.awt.MessageBoxButtons import (BUTTONS_YES_NO,
                                                    BUTTONS_OK, )


class MessageBoxResults:
    # noinspection PyUnresolvedReferences
    from com.sun.star.awt.MessageBoxResults import (
        OK, CANCEL, YES)


def get_collator_from_doc(
        ctxt: XComponentContext, oDoc: XPropertySet, ignore_case: bool = True
) -> XCollator:
    oServiceManager = ctxt.ServiceManager
    oCollator = oServiceManager.createInstance(
        "com.sun.star.i18n.Collator")

    oLocale = oDoc.CharLocale
    oCollator.loadDefaultCollator(oLocale, 1 if ignore_case else 0)
    return oCollator


def get_whole_cell(ctxt: XComponentContext) -> bool:
    node_path = "org.openoffice.Office.Calc/Calculate/Other"
    oAccess = _access_node(ctxt, node_path)
    return oAccess.SearchCriteria


def _access_node(ctxt: XComponentContext, node_path: str) -> Any:
    pv = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    pv.Name = "nodepath"
    pv.Value = node_path
    oConfigProvider = ctxt.getValueByName(
        "/singletons/com.sun.star.configuration.theDefaultProvider")
    oAccess = oConfigProvider.createInstanceWithArguments(
        "com.sun.star.configuration.ConfigurationAccess", (pv,))
    return oAccess


NAMES_BY_VERSION = {
    "24.8": [
        "FILTER", "RANDARRAY", "SEQUENCE", "SORTBY", "SORT", "UNIQUE",
        "XLOOKUP", "XMATCH"
    ],
    "25.8": [
        "CHOOSECOLS", "CHOOSEROWS", "DROP", "TAKE", "EXPAND", "HSTACK",
        "VSTACK", "TOCOL", "TOROW", "WRAPCOLS", "WRAPROWS",
    ]
}


def upgrade(ctxt: XComponentContext, oDoc) -> Any:
    import logging
    logger = logging.getLogger(__name__)
    logger.debug("Upgrade")

    oStatusIndicator = oDoc.CurrentController.StatusIndicator
    oStatusIndicator.start("Upgrade", 100)

    try:
        lo_version = _access_node(
            ctxt, "org.openoffice.Setup/Product").ooSetupVersion
        all_names = []
        for version, names in NAMES_BY_VERSION.items():
            if lo_version >= version:
                all_names.extend(names)

        oServiceManager = ctxt.ServiceManager
        oToolkit = oServiceManager.createInstance("com.sun.star.awt.Toolkit")

        if not all_names:
            oMessageBox = oToolkit.createMessageBox(
                None, MessageBoxType.MESSAGEBOX,
                MessageBoxButtons.BUTTONS_OK,
                "Upgrade", "\n".join([
                    "Detected LibreOffice Version: {}".format(lo_version),
                    "No LOP function to upgrade"
                ])
            )
            oMessageBox.execute()
            return "No upgrade"

        oMessageBox = oToolkit.createMessageBox(
            None, MessageBoxType.MESSAGEBOX, MessageBoxButtons.BUTTONS_YES_NO,
            "Upgrade", "\n".join([
                "Detected LibreOffice Version: {}".format(lo_version),
                "Upgrade {} to LibreOffice standard?".format(
                    ", ".join(["LOP.{}".format(n) for n in all_names]))
            ])
        )
        if oMessageBox.execute() != MessageBoxResults.YES:
            return "No upgrade"

        regex = re.compile(
            r"COM\.GITHUB\.JFERARD\.LOPOLYFILL\.LOPOLYFILLIMPL\.LOP({})".format(
                "|".join(all_names)))

        oSheets = oDoc.Sheets

        def func(m: re.Match) -> str:
            return m.group(1)

        count = 0
        for i in range(oSheets.Count):
            oSheet = oSheets.getByIndex(i)
            logger.debug("Upgrade sheet %s", oSheet.Name)
            oStatusIndicator.Text = oSheet.Name
            oStatusIndicator.Value = (i * 100) // oSheets.Count

            oCellRanges = oSheet.queryContentCells(16)
            for j in range(oCellRanges.Count):
                oCellRange = oCellRanges.getByIndex(j)
                oCell = oCellRange.getCellByPosition(0, 0)
                if oCellRange.ArrayFormula != "":
                    clean_formula = oCell.Formula[2:-1]
                    new_clean_formula = regex.sub(func, clean_formula)
                    if new_clean_formula != clean_formula:
                        oCellRange.ArrayFormula = clean_formula
                        count += 1
                else:
                    formula = oCell.Formula
                    new_formula = regex.sub(func, formula)
                    if new_formula != formula:
                        oCell.Formula = new_formula
                        count += 1

        msg = "Upgraded {} formula(s)".format(count)
        oMessageBox = oToolkit.createMessageBox(
            None, MessageBoxType.MESSAGEBOX,
            MessageBoxButtons.BUTTONS_OK,
            "Upgrade", msg
        )
        oMessageBox.execute()
        return msg
    except Exception as _e:
        logger.exception("Upgrade")
    finally:
        oStatusIndicator.end()
