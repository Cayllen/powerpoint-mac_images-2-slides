# Images-2-Powerpoint-Slides

### Why does this tool exist?
- Powerpoint on Mac does not support to load multiple images on individual powerpoint slides - thus that has to be done manually

###What does the tool do?
- This simple Python tool asks you to select all photos that should be placed in a Powerpoint.
- After that each image gets scaled up, so it fits on the slide
    - Image horizontal: Scales to the full slide-size (no whitespace on the top; may cut off sides)
    - Image vertical: Scales so that the full image is visible (with whitespace left + right)
- The script tries to check if the image is rotated