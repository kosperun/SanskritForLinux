#!/bin/bash
# Installs Sanskrit conversion macros into the LibreOffice Standard library.
# Run this script with LibreOffice closed.
# After running, restart LibreOffice and the macros will be available under
# Tools -> Macros -> Run Macro -> My Macros -> Standard -> SanskritConversions.

set -e

MACRO_DIR="$(cd "$(dirname "$0")/LibreOffice Macros" && pwd)"

if [[ "$OSTYPE" == "darwin"* ]]; then
    LO_BASIC="$HOME/Library/Application Support/LibreOffice/4/user/basic/Standard"
elif [[ "$OSTYPE" == "linux"* ]]; then
    LO_BASIC="$HOME/.config/libreoffice/4/user/basic/Standard"
else
    echo "Unsupported OS: $OSTYPE"
    exit 1
fi

if [ ! -d "$LO_BASIC" ]; then
    echo "LibreOffice user profile not found at: $LO_BASIC"
    echo "Make sure LibreOffice is installed and has been opened at least once."
    exit 1
fi

echo "Installing to: $LO_BASIC"

MODULE_NAME="SanskritConversions"
{
    echo '<?xml version="1.0" encoding="UTF-8"?>'
    echo '<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">'
    printf '<script:module xmlns:script="http://openoffice.org/2000/script" script:name="%s" script:language="StarBasic" script:moduleType="normal"><![CDATA[\n' "$MODULE_NAME"
    cat "$MACRO_DIR/SanskritConversions.bas"
    printf '\n]]></script:module>\n'
} > "$LO_BASIC/$MODULE_NAME.xba"
echo "  Installed: $MODULE_NAME.xba"

# Remove old individual module files if present
for old in ConversionHelpers Balaram_to_NoDiacritics Balaram_to_Russian_Cyrillic \
           Balaram_to_Ukrainian_Cyrillic Balaram_to_Unicode_LO_macros \
           Cyrillic_Unicode_to_Balaram_LO_macros HarvardKyoto_to_Russian_Cyrillic \
           HarvardKyoto_to_Ukrainian_Cyrillic UnicodeToBalaram_LO_marcos TestSelection; do
    rm -f "$LO_BASIC/$old" "$LO_BASIC/$old.xba"
done

# Rewrite script.xlb preserving existing user modules
XLB="$LO_BASIC/script.xlb"
existing_entries=""
if [ -f "$XLB" ]; then
    while IFS= read -r line; do
        if [[ "$line" == *"library:element"* ]]; then
            entry_name=$(echo "$line" | sed 's/.*library:name="\([^"]*\)".*/\1/')
            if [ "$entry_name" != "$MODULE_NAME" ] && \
               [ "$entry_name" != "ConversionHelpers" ] && \
               [ "$entry_name" != "TestSelection" ] && \
               [[ "$entry_name" != "Balaram_"* ]] && \
               [[ "$entry_name" != "Cyrillic_"* ]] && \
               [[ "$entry_name" != "HarvardKyoto_"* ]] && \
               [[ "$entry_name" != "UnicodeToBalaram_"* ]]; then
                existing_entries="$existing_entries\n $line"
            fi
        fi
    done < "$XLB"
fi

{
    echo '<?xml version="1.0" encoding="UTF-8"?>'
    echo '<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">'
    echo '<library:library xmlns:library="http://openoffice.org/2000/library" library:name="Standard" library:readonly="false" library:passwordprotected="false">'
    if [ -n "$existing_entries" ]; then
        echo -e "$existing_entries"
    fi
    echo " <library:element library:name=\"$MODULE_NAME\"/>"
    echo '</library:library>'
} > "$XLB"

echo ""
echo "Done. Restart LibreOffice to use the macros."
echo "Find them under: Tools -> Macros -> Run Macro -> My Macros -> Standard -> SanskritConversions"
