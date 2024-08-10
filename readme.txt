# Radiology Right Click v1.0

## Description
Radiology Right Click v1.0 is a radiologist's best friend, designed to improve quantitative radiology and streamline workflow. This AutoHotkey script provides a suite of tools accessible via right-click, offering quick calculations and reference information for various radiological measurements and assessments.

It works by highlighting text in PowerScribe360 as installed at the authors institution or NotePad for troubleshooting.

The user highlights text and right clicks with the mouse. A new right click menu appears with various useful functions. 
The script can be paused to see the default right click menu for a specified time - 3 minutes default and up to 10 hours to turn off for an entire shift. 

## Author
Created by an amateur programming enthusiast with a passion for radiology, in partnership with Claude, an AI assistant.

## Features
- Calcium Score Percentile Calculator (MESA and Hoff methods)
- Ellipsoid Volume Calculator
- PSA Density Calculator
- Pregnancy Date Calculator
- Menstrual Phase Estimator
- Adrenal Washout Calculator
- Thymus Chemical Shift Calculator
- Hepatic Steatosis Estimator
- MRI Liver Iron Content Estimator
- Nodule Size Comparison and Sorting (Sorting takes in 3 comma separated numbers and returns in descending order separated by x to improve formatting of ModLink fields containing 3 comma separated values).
- Number statistics
- Number formatting correction 



## Usage

1) Download script
2) Download AHK converter and compile script. This may need to be done on a personal PC. 
2a) Optional: Place the generated exe file on a usb stick.
2b) Optional: Plug usb stick into PACS work station
3) Start the script by double clicking the exe file.
4) Highlight text to process in PowerScribe (or NotePad in test environment) 
5) Right-click on mouse, selected calculation menu choice.
6) To turn off the script, click "Show Hidden Icons" on Windows taskbar (^ carrot), Find the Green letter H indicating an AutoHotKey script is running, right click on the script and select "exit"

Troubleshooting: Ensure you are only running one instance of the script if there is unexpected behavior by going to the "Show Hidden Icons" in the Windows Task Bar and making sure only 1 green letter H is present. Exit duplicate instances as needed.




## Status
This project is currently in testing and development. We are aware of several areas for improvement:
- Global variables need to be cleaned up (removed) for better code organization
- Local machine learning functions could be implemented for improved text processing and variable extraction
- Potential for automatic impression generation based on findings

## Future Improvements
- Expand compatibility beyond PowerScribe to support more radiology reporting software (may be just a simple variable definition and preference setting based on in Windows active programs/processes)
- Enhance the preferences menu for better user customization
- Add support for additional radiology calculations as packages, including:
  - DEXA analysis
  - Custom MR or US elastography equations
  - MSK-specific calculations
  - Cardiac imaging metrics
  - Neuro-related assessments
- Improve the modularity of the code to allow for easier addition of new calculation packages
- Develop functionality to process entire reports for overall quantitative change in mass/nodule/lesion size:
  - Implement intelligent exclusion of measurements for collections
  - Incorporate advanced topics such as RECIST (Response Evaluation Criteria in Solid Tumors)
  - Note: Full implementation may require NLP or ML packages, which presents challenges in hospital settings due to security restrictions on external server connections
- Explore options for secure, locally-run NLP and ML capabilities to enable advanced text processing without compromising hospital network security
- Implement spelling and grammar checking capabilities:
  - Detect and suggest corrections for common radiological terms and phrases
  - Identify potential dictation errors and offer corrections
  - Ensure consistency in terminology throughout the report
- Develop context-aware error detection:
  - Identify discrepancies between measurements in the body and impression of the report
  - Flag potential contradictions or inconsistencies in the report
- Create a customizable medical abbreviation expander:
  - Allow users to define and expand commonly used abbreviations in their reports
  - Implement smart expansion based on context to avoid errors
- Add support for nuclear medicine calculations:
  - Tracer dosage calculations based on patient weight, age, and specific study type
  - Decay calculations for various radioisotopes
  - SUV (Standardized Uptake Value) calculations for PET studies
  - Glomerular Filtration Rate (GFR) estimations for renal studies
  - Thyroid uptake calculations
  - Gastric emptying time calculations
- Implement a module for radiation safety calculations:
  - Effective dose estimations for different imaging modalities
  - Cumulative radiation dose tracking for patients across multiple studies
- Develop modular decision support algorithms and radiology reporting tools:
  - BI-RADS (Breast Imaging-Reporting and Data System) for mammography, ultrasound, and MRI
  - TI-RADS (Thyroid Imaging Reporting and Data System)
  - LI-RADS (Liver Imaging Reporting and Data System)
  - O-RADS (Ovarian-Adnexal Imaging Reporting and Data System)
  - Lung-RADS (Lung CT Screening Reporting and Data System)
  - PI-RADS (Prostate Imaging-Reporting and Data System)
  - C-RADS (CT Colonography Reporting and Data System)
  - NI-RADS (Neck Imaging Reporting and Data System)
  - HI-RADS (Head Injury Imaging Reporting and Data System)
- Implement interactive decision trees for each RADS system to guide radiologists through the classification process
- Create customizable reporting templates based on RADS classifications
- Develop a feature for automatic generation of follow-up recommendations based on RADS assessments
- Add multi-language support:
  - Implement a localization system for the user interface
  - Provide translations for all calculations, reports, and decision support tools
  - Include support for right-to-left languages
  - Develop a system for easy addition of new languages by contributors
  - Ensure compatibility with various character sets and input methods
  - Implement language-specific formatting for numbers, dates, and units
  - Create multilingual templates for structured reporting
- Develop automated measurement capture and reporting system:
  - Implement PACS integration to automatically capture measurements made on images
  - Create a feature to automatically populate report fields with captured measurements
  - Develop intelligent field mapping to place measurements in appropriate report sections
  - Implement automatic tabbing to the next field after populating a measurement
  - Add support for multiple measurement series and labeling (e.g., "Lesion 1", "Lesion 2")
  - Develop a user interface for mapping PACS measurement labels to report fields
  - Implement error checking to flag significant discrepancies between PACS measurements and manually entered values
  - Create an undo/redo system for measurement insertions
  - Develop a feature to manually trigger measurement insertion for specific fields
  - Implement support for various PACS systems and reporting software to ensure broad compatibility
- Implement auto-population of fields based on technologist notes:
  - Develop integration with PACS and electronic health record (EHR) systems to access technologist notes
  - Create an intelligent parsing system to extract relevant information from technologist notes
  - Implement automatic field population based on extracted information
  - Develop a mapping system to match extracted information with appropriate report fields
  - Create a user interface for customizing the mapping between note content and report fields
  - Implement a highlighting system to indicate auto-populated fields for easy review
  - Develop a feature to compare auto-populated information with PACS measurements for consistency
  - Create an override system allowing radiologists to easily modify or reject auto-populated content
  - Implement support for various PACS and EHR systems to ensure broad compatibility
  - Develop a learning system to improve accuracy of information extraction over time based on radiologist corrections





## Sample Inputs

### Calcium Score Percentile
```
Age: 56
Sex: Male
Race: White (this is optional and leads to an external HTTP call for MESA numbers)

YOUR CORONARY ARTERY CALCIUM SCORE: 56 (Agatston)
```

### Ellipsoid Volume
```
3.5 x 2.6 x 4.1 cm
```

### PSA Density
```
PSA level: 4.5 ng/mL, Prostate volume: 30 cc
```

### Pregnancy Dates
```
LMP: 01/15/2023
```
or
```
GA: 12 weeks and 3 days as of today
```
OR
```
GA: 12 weeks and 3 days as of 01/01/2021
```

### Menstrual Phase
```
LMP: 05/01/2023
```

### Adrenal Washout
```
Unenhanced: 10 HU, Enhanced: 80 HU, Delayed: 40 HU
```

### Thymus Chemical Shift
```
Thymus IP: 100, OP: 80, Paraspinous IP: 90, OP: 85
```

### Hepatic Steatosis
```
Liver IP: 100, OP: 80, Spleen IP: 90, OP: 88
```

### MRI Liver Iron Content
```
1.5T, R2*: 50 Hz
```

### Nodule Size Comparison (takes in 1, 2 or 3 measurements)
```
previous 3.5 x 2.6 cm  now 4.1 x 2.9 cm
OR 
```
4.1 x 2.9 x 3.3 cm (previously 3.9 x 2.8 x 3.0 cm)
```

### Sort measurement sizes (takes in 2 or 3 measurements)
```
3.5 x 2.6 x 4.2 cm -->  4.2 x 3.5 x 2.6 cm 
```
OR
```
(3,4,7) --> 7 x 4 x 3

### Calculate number range (takes in any number of digits)
```
3,12,5 --> 3 - 12
```
OR
3.5 4.2 5.5 --> 3.5 - 5.5

### Calculate statistics (takes in any number of digits)
```
Returns various statistics meaningful to daily practice. More statistics are returned for larger datasets.


## Installation
1. Install AutoHotkey v1.1
2. Download the script file
3. Run the script

## Contributions
Contributions, suggestions, and bug reports are welcome. Please open an issue or submit a pull request on GitHub. We're particularly interested in contributions that address our planned future improvements or add new calculation packages for various radiology subspecialties, including nuclear medicine. We also welcome innovative solutions for implementing advanced text processing capabilities within the constraints of hospital security protocols, including spelling and grammar checking, dictation error detection, and medical abbreviation expansion. Contributions to the development of RADS modules and decision support algorithms are highly encouraged. We are also seeking contributors for translations and localization to expand the tool's accessibility to non-English speaking radiologists. Expertise in PACS integration, EHR systems, automated data capture, and natural language processing would be especially valuable for implementing the automated measurement capture and note parsing systems.

## License

Radiology Right Click v1.0 is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.

## Disclaimer
This software is provided for educational and research purposes only. It should not be used as a substitute for professional medical advice, diagnosis, or treatment. Always seek the advice of your physician or other qualified health provider with any questions you may have regarding a medical condition. Users should exercise caution and verify all calculations, especially those related to radiation dosage and patient safety. While the RADS modules aim to assist in classification and reporting, the final interpretation and clinical decisions should always be made by qualified healthcare professionals. Users should be aware that translations and localizations may not be officially certified for medical use in all jurisdictions. The automated measurement capture and note parsing features are designed to assist radiologists but should not replace careful review and validation of all information in the final report. Radiologists are responsible for ensuring the accuracy and appropriateness of all auto-populated content.
