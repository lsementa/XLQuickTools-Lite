// --- Menu Control Functions ---
// Menu.js


Office.onReady(() => {
    // On Ready
});

// Get a reference to your button element
const undoButton = document.getElementById("undoButton");

// Function to disable the undo button
function disableUndoButton() {
    document.getElementById('undoButton').classList.add('disabled');
    // Hide the header dropdown
    const dropdown = document.querySelector('.header-dropdown-menu');
    if (dropdown) dropdown.classList.remove('show');
}

// Function to enable the undo button
function enableUndoButton() {
    document.getElementById('undoButton').classList.remove('disabled');
}

// Toggle the visibility of the dropdown menu
function toggleDropdown(element, event) {
    let dropdownMenu;

    if (element.classList.contains('dropdown-toggle')) {
        // Split button (toggle part)
        dropdownMenu = element.nextElementSibling;

    } else if (element.classList.contains('full-dropdown-button')) {
        // Full dropdown button
        dropdownMenu = element.querySelector(".dropdown-menu");

    } else if (element.classList.contains('header-button')) {
        // Header button dropdown
        dropdownMenu = element.nextElementSibling;

    } else {
        console.error("toggleDropdown called on an unexpected element:", element);
        return;
    }

    if (!dropdownMenu) return;

    // Hide all other dropdowns (including header)
    document.querySelectorAll(".dropdown-menu.show, .header-dropdown-menu.show").forEach(menu => {
        if (menu !== dropdownMenu) {
            menu.classList.remove("show");
        }
    });

    // Toggle the current menu
    dropdownMenu.classList.toggle("show");
}

// Close all dropdowns when clicking outside
document.addEventListener('click', function (event) {
    const isClickInsideDropdown =
        event.target.closest('.split-button') ||
        event.target.closest('.full-dropdown-button') ||
        event.target.closest('.header');

    if (!isClickInsideDropdown) {
        document.querySelectorAll(".dropdown-menu.show, .header-dropdown-menu.show").forEach(menu => {
            menu.classList.remove("show");
        });
    }
});


// Menu item clicked
async function handleClick(action, event) {
    // console.log("Action (from Task Pane):", action);

    try {
        switch (action) {
            case 'home':
                window.location.href = 'Menu.html';
                break;
            case 'help':
                window.location.href = 'Help.html';
                break;
            case 'about':
                window.location.href = 'About.html';
                break;
            case 'undo':
                if (undoManager.canUndo()) {
                    await undoManager.undoFormatting();
                    // Disable after successful undo
                    disableUndoButton();
                }
                break;
            case 'trim-clean-selected':
                trimCleanSelected();
                break;
            case 'trim-clean-sheet':
                trimCleanSheet();
                break;
            case 'trim-clean-workbook':
                trimCleanWorkbook();
                break;
            case 'remove-excess':
                removeExcess();
                break;
            case 'text-uppercase':
                getTextOptions(TextTransformOption.UPPERCASE);
                break;
            case 'text-lowercase':
                getTextOptions(TextTransformOption.LOWERCASE);
                break;
            case 'text-propercase':
                getTextOptions(TextTransformOption.PROPERCASE);
                break;
            case 'text-remove-letters':
                getTextOptions(TextTransformOption.REMOVE_LETTERS);
                break;
            case 'text-remove-numbers':
                getTextOptions(TextTransformOption.REMOVE_NUMBERS);
                break;
            case 'text-remove-special':
                getTextOptions(TextTransformOption.REMOVE_SPECIAL);
                break;
            case 'subscript-numbers':
                getTextOptions(TextTransformOption.SUBSCRIPT_UNICODE);
                break;
            case 'text-add-leadtrail':
                showAddLeadTrailModal();
                break;
            case 'date-converter':
                showDateTextModal();
                break;
            case 'delete-empty-rows':
                deleteEmptyRows();
                break;
            case 'delete-empty-columns':
                deleteEmptyColumns();
                break;
            case 'remove-hyperlinks':
                removeHyperlinks();
                break;
            case 'add-hyperlinks':
                showAddHyperlinksModal();
                break;
            case 'fill-down':
                fillBlanksFromAbove();
                break;
            case 'split-to-rows':
                showSplitToRowsModal();
                break;
            case 'selection-to-clipboard':
                showSelectionPlusModal();
                break;
            case 'check-duplicates':
                findAndCountDuplicatesInColumn();
                break;
            case 'toggle-autofilter':
                toggleAutoFilter();
                break;
            case 'copy-highlighted':
                copyHighlightedClipboard();
                break;
            case 'unique-count':
                countUniqueValuesInColumn();
                break;
            case 'find-missing-data':
                showFindMissingModal();
                break;
            case 'compare-worksheets':
                showCompareSheetsModal();
                break;
            case 'reset-column':
                resetColumn();
                break;
            default:
                console.warn("Menu.js: Unknown action: " + action);
        }

        // Hide dropdown after run
        if (this && this.classList && this.classList.remove) {
            this.classList.remove("show");
        }

    } catch (error) {
        console.error("Menu.js: Error during Excel operation:", error);
    }
}

// Event listeners
window.handleClick = handleClick;
window.toggleDropdown = toggleDropdown;

