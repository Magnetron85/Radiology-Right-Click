# Radiology Right Click v1.04

(c) 2024, GitHub user Magnetron85

## Description

Radiology Right Click v1.04 is a radiologist's quantitative companion, designed to enhance quantitative radiology and streamline workflow. This AutoHotkey script provides a suite of tools accessible via right-click in PowerScribe at the point of dictation, offering quick calculations and reference information for various radiological measurements and assessments. The script is easily modifiable for non-PowerScribe environments.

Key features:
- Works by highlighting text in PowerScribe360 or NotePad (for troubleshooting)
- Customizable right-click menu with various useful functions
- Pause feature to access the default right-click menu
- User-configurable calculator visibility
- Optional OCR for coronary calcium scoring (requires Vis2 OCR for AHK)

## Author

Created by GitHub user Magnetron85, an amateur programming enthusiast with a passion for radiology, in collaboration with Claude, an AI assistant.

Contributions for Fleischner module by Eric Van Bogaert (https://github.com/EricVanBogaert/AutoHotKey-for-Radiology)

## Features and Usage

Below are detailed descriptions of each available function, along with example inputs and notes on edge cases:

1. **Coronary Artery Calcium Score Percentile Calculator (MESA and Hoff methods)**
   Description: Calculates age and sex-matched coronary artery calcification percentile.
   Example input:
   ```
   Age: 56
   Sex: Male
   Race: White
   YOUR CORONARY ARTERY CALCIUM SCORE: 56 (Agatston)
   ```
   Notes: 
   - Race is optional; if omitted, Hoff method is used.
   - Supports ages 45-84 for MESA, 30-74 for Hoff. If age < 45, Hoff is used even if race is provided

2. **Ellipsoid Volume Calculator**
   Description: Calculates volume of an ellipsoid based on three dimensions.
   Example input:
   ```
   3.5 x 2.6 x 4.1 cm
   ```
   Notes:
   - Accepts measurements in cm or mm.
   - Assumes cm if decimal points are present, mm otherwise.

3. **Bullet Volume Calculator**
   Description: Calculates volume using the bullet method.
   Example input:
   ```
   3.5 x 2.6 x 4.1 cm
   ```
   Notes:
   - Input format same as Ellipsoid Volume Calculator.

4. **PSA Density Calculator**
   Description: Calculates PSA density given PSA level and prostate volume or dimensions.
   Example inputs:
   ```
   PSA level: 4.5 ng/mL, Prostate volume: 30 cc
   ```
   or
   ```
   PSA level: 4.5 ng/ml, Prostate Size: 4.5 x 4.2 x 4.1 cm
   ```
   Notes:
   - Can calculate volume if dimensions are provided.
   - Switches to Ellipsoid volume if calculated Bullet volume > 55 cc.

5. **Pregnancy Date Calculator**
   Description: Calculates LMP, current gestational age, and expected delivery date.
   Example inputs:
   ```
   LMP: 01/15/2023
   ```
   or
   ```
   GA: 12 weeks and 3 days as of today
   ```
   or
   ```
   GA: 12 weeks and 3 days as of 01/01/2021
   ```
   Notes:
   - Accepts various date formats and gestational age inputs.

6. **Menstrual Phase Estimator**
   Description: Calculates the current date of menstrual cycle based on a 28-day cycle.
   Example input:
   ```
   LMP: 05/01/2023
   ```

7. **Adrenal Washout Calculator**
   Description: Calculates absolute and relative washout for adrenal lesions.
   Example input:
   ```
   Unenhanced: 10 HU, Enhanced: 80 HU, Delayed: 40 HU
   ```
   Notes:
   - Can calculate relative washout if only enhanced and delayed values are provided.

8. **Thymus Chemical Shift Calculator**
   Description: Calculates thymic chemical shift ratio and/or signal intensity index.
   Example input:
   ```
   Thymus IP: 100, OP: 80, Paraspinous IP: 90, OP: 85
   ```
   Notes:
   - Can calculate signal intensity index if only thymus values are provided.

9. **Hepatic Steatosis Estimator**
   Description: Estimates liver fat fraction based on in-phase and out-of-phase MRI.
   Example input:
   ```
   Liver IP: 100, OP: 80, Spleen IP: 90, OP: 88
   ```
   Notes:
   - Can calculate fat fraction without spleen values, but less accurate.

10. **MRI Liver Iron Content Estimator**
    Description: Estimates liver iron content based on R2* values.
    Example input:
    ```
    1.5T, R2*: 50 Hz
    ```
    Notes:
    - Supports 1.5T and 3.0T field strengths.

11. **Nodule Size Comparison and Sorting**
    Description: Compares nodule sizes over time and sorts measurements.
    Example inputs:
    ```
    4.1 x 2.9 x 3.3 cm (previously 3.9 x 2.8 x 3.0 cm)
    ```
    or
    ```
    4.1 cm, previously 3.9 cm on 01/01/2000
    ```
    Notes:
    - Accepts various input formats, including dates if provided.
    - Calculates volume change, doubling time, and growth rate when possible.

12. **Number Statistics**
    Description: Calculates various statistics for a set of numbers.
    Example input:
    ```
    3, 12, 5, 8, 15, 7
    ```
    Notes:
    - Provides more detailed statistics for larger datasets (9+ numbers).

13. **Measurement Formatting Correction**
    Description: Corrects and standardizes measurement formatting.
    Example input:
    ```
    3.5 x 2.6 x 4.2 cm
    ```
    Notes:
    - Sorts measurements in descending order.
    - Handles multiple sets of measurements in a single input.

14. **OCR for Coronary Calcium Scoring (BETA)**
    Description: Captures and processes coronary calcium scores from images.
    Notes:
    - Requires Vis2 OCR for AHK.
    - Tested with syngo.via.
    - Attempts to correct for common OCR errors.
    - User can define the coordinates of their syngo.via table (x and y pixel values at top left of table) and size (width and height of table in pixel values) in preferences. 
    - Note: leave monitor set to 1 if coordinates are global (this is often the case if using Windows Spy to get the coordinate).

15. **Fleischner 2017 Recommendations**
    Description: Provides Fleischner Society 2017 recommendations for pulmonary nodules.
    Example input:
    ```
    Incidental right upper lobe solid noncalcified pulmonary nodule measuring 7 x 8 mm (series 1, image 30).
    ```
    Notes:
    - Handles multiple nodules and various descriptors (solid, ground-glass, part-solid).
    - Accounts for patient risk factors.
    - Provides follow-up dates based on current date.

16. **Contrast Premedication Calculator**
    Description: Calculates contrast premedication schedule based on scan time.
    Notes:
    - Provides options for prednisone and methylprednisolone protocols.
    - Includes optional diphenhydramine dosing.

17. **NASCET Calculator**
    Description: Calculates the North American Symptomatic Carotid Endarterectomy Trial (NASCET) value for carotid artery stenosis.
    Example inputs:
    ```
    Distal ICA: 5mm, stenosis: 2mm
    5 mm, 0.2 cm
    The narrowest lumen measures 2 mm. The distal lumen measures 5 mm.
    The tightest lumen measures 0.3 cm. The patent lumen measures 9 mm.
    ```
    Notes: 
    - Accepts measurements in mm or cm.
    - Can handle various descriptive terms for stenosis (e.g., "narrowest", "tightest") and normal lumen (e.g., "distal", "patent").
    - Intelligently infers which measurement is the stenosis (smaller) and which is the distal ICA (larger) when labels are unclear.
    - Provides NASCET percentage and interpretation (mild, moderate, or severe stenosis).
    - Includes citation for the NASCET study when citations are enabled.

## Installation and Usage

The script does not require installation on a hospital-owned workstation. However, you may or may not have permission to download .exe files.

### Pathway #1: Using the pre-compiled executable

1. Download the .exe file from GitHub.
2. Download the OCR dependency (https://github.com/iseahound/Vis2).
3. Place both files in the same directory.
4. Double-click the .exe file to run.

If you do not have permission to download .exe files:
1. Download on a personal computer and place on a USB stick.
2. Plug the USB stick into your work computer, then double-click and run.

### Pathway #2: Compiling from source

1. Download AutoHotKey v1.1 on a PC where you have permission to install software (often a personal PC).
2. Download the script from GitHub on the same PC where you installed AutoHotKey v1.1.
3. Download the OCR dependency (https://github.com/iseahound/Vis2) and place its files in the same directory as the script.
4. Compile the script using the AutoHotKey v1.1 Compiler.
5. (Optional) Place the generated .exe file on a USB stick to transfer to your PACS workstation. Include the dependency folders in the same directory on the USB stick.
6. Plug the USB stick into your PACS workstation (if using this method).
7. Start the script by double-clicking the .exe file generated by the AutoHotKey v1.1 Compiler.

### Usage

1. Highlight text to process in PowerScribe360 (or NotePad in test environment).
2. Right-click with the mouse.
3. Select a menu choice from the new right-click menu that appears.
4. Functions return results in message boxes with output also pasted to your clipboard. Input format is (usually) preserved on the clipboard to allow you to directly paste into reports with minimal editing.
5. To turn off the script, click "Show Hidden Icons" on the Windows taskbar (^ carrot), find the Green letter H indicating an AutoHotKey script is running, right-click on the script and select "exit".

## Troubleshooting

1. Ensure you are only running one instance of the script if there is unexpected behavior:
   - Go to the "Show Hidden Icons" in the Windows Task Bar.
   - Make sure only 1 green letter H is present.
   - Exit duplicate instances as needed.

2. If the script is not responding to right-clicks:
   - Check if the script is running (look for the green H icon in the system tray).
   - Ensure you're using it in a supported application (PowerScribe360 or Notepad).
   - Try restarting the script.

3. If OCR for calcium scoring is not working:
   - Verify that the Vis2 OCR dependency is in the correct directory.
   
4. If calculations seem incorrect:
   - Double-check your input format against the examples provided in the README.
   - Remember that some functions make assumptions (e.g., cm vs mm) based on the presence of decimal points.

5. If the script is conflicting with other software:
   - Use the pause feature to temporarily disable the script.
   - Consider adjusting the script's target applications in the code.

6. If you're getting unexpected results from the Fleischner recommendations:
   - Ensure your input includes all relevant information (size, composition, patient risk factors).
   - Check that the nodule description is clear and follows the expected format.

## Contributions

We welcome contributions, suggestions, and bug reports. Here are some specific areas where we're looking for help:

1. **New Calculation Packages**: Develop new modules for various radiology subspecialties, including nuclear medicine.

2. **Advanced Text Processing**: Implement innovative solutions for text processing within hospital security constraints. This could include:
   - Spelling and grammar checking
   - Dictation error detection
   - Medical abbreviation expansion

3. **RADS Modules**: Develop decision support algorithms for various RADS (Reporting and Data Systems) used in radiology.

4. **Language Support**: Help with translations and localization to make the tool accessible to non-English speaking radiologists.

5. **Integration**: Contribute expertise in PACS integration and EHR systems to improve the tool's interoperability.

6. **Automated Data Capture**: Implement systems for automated measurement capture from PACS or reports.

7. **Natural Language Processing**: Develop more advanced NLP capabilities for parsing radiology reports and extracting relevant information.

8. **User Interface Improvements**: Enhance the tool's UI for better user experience, potentially moving beyond AutoHotkey's limitations.

9. **Documentation**: Improve and expand the documentation, including creating user guides and API documentation for developers.

10. **Testing and Quality Assurance**: Help develop a comprehensive testing suite to ensure the tool's accuracy and reliability.

To contribute:

1. Fork the repository on GitHub.
2. Create a new branch for your feature or bug fix.
3. Commit your changes with clear, descriptive commit messages.
4. Push your branch and submit a pull request with a clear description of your changes.

Please ensure your code adheres to the existing style for consistency. For major changes, please open an issue first to discuss what you would like to change.

We also appreciate bug reports and feature requests. Please use the GitHub issue tracker for these.
## License

Radiology Right Click v1.04 is licensed under the GNU General Public License v3.0. See the LICENSE file for full details.

## Disclaimer

This software is provided for educational and research purposes only. It should not be used as a substitute for professional medical advice, diagnosis, or treatment. Always seek the advice of your physician or other qualified health provider with any questions you may have regarding a medical condition. Users should exercise caution and verify all calculations for patient safety. The developers and software are not responsible for clinical decisions or errors resulting from the use of this tool.
