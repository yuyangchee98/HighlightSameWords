# Highlight same words

A Microsoft Word VBA macro. When you click on a word and run the macro, it highlights all instances of that word in the document. Click again to remove the highlighting.

![Demo](https://i.imgur.com/anCrk85.gif)

## Features

- Highlight all instances of a word by clicking on it
- Works with selected phrases (multiple words)
- Toggle highlighting on/off with second click
- Yellow highlighting (can be modified in the code)
- Matches whole words only to avoid partial matches

## Installation

1. Open Microsoft Word
2. Press Alt + F11 to open the Visual Basic Editor
3. In the Project Explorer (left side), right-click "Modules"
4. Select Insert > Module
5. Copy and paste the contents of `HighlightSameWords.vba` into the new module
6. Save and close the VBA editor
7. Add to the ribbon or as a keyboard shortcut as you please

## Usage

1. Click anywhere in a word or select multiple words
2. Either:
   - Run the macro from the Macros menu (View > Macros)
   - Or press your assigned keyboard shortcut
3. Click the same word and run the macro again to remove highlighting
