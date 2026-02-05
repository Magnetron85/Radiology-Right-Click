# Radiology Right Click v1.22 and v1.22 local

A radiologist's quantitative companion designed to enhance workflow through right-click calculations, reference tools and collecting commonly used LLM prompts. This AutoHotkey script provides quick access to common radiological measurements and assessments directly in PowerScribe or other supported environments.

## What's New in v1.22

- Added customizable menu sorting
- Improved multi-monitor support 
- New references section for saving URLs and files
- Added a spell check that looks up highlighted words with google 
- Added experimental AI integration with Claude/ChatGPT prompt templates 
- v1.22 local does not include external web calls (No MESA website lookup, no AI option, no google spell check). Do not ever send PHI to a third party. Dont use the experimental component with real patient information.


## Features

- Works by highlighting text in PowerScribe360 or NotePad
- Customizable right-click menu with various calculations
- Pause feature to access default right-click menu
- User-configurable calculator visibility
- Custom menu sorting options
- References management for quick access to resources
- AI assistant integration for report analysis, proof of concept, Not to be used clinically due to privacy.

## Calculations & Tools

1. **Coronary Artery Calcium Score Percentile (MESA/Hoff in v1.22 and Hoff in v1.22-local)**
   ```
   Age: 56
   Sex: Male
   Race: White
   YOUR CORONARY ARTERY CALCIUM SCORE: 56 (Agatston)
   ```

2. **Ellipsoid Volume**
   ```
   3.5 x 2.6 x 4.1 cm
   ```

3. **Bullet Volume**
   ```
   3.5 x 2.6 x 4.1 cm
   ```

4. **PSA Density**
   ```
   PSA level: 4.5 ng/mL, Prostate volume: 30 cc
   ```
   or
   ```
   PSA level: 4.5 ng/ml, Prostate Size: 4.5 x 4.2 x 4.1 cm
   ```

5. **Pregnancy Dates**
   ```
   LMP: 01/15/2023
   ```
   or
   ```
   GA: 12 weeks and 3 days as of today
   ```

6. **Menstrual Phase**
   ```
   LMP: 05/01/2023
   ```

7. **Adrenal Washout**
   ```
   Unenhanced: 10 HU, Enhanced: 80 HU, Delayed: 40 HU
   ```

8. **Thymus Chemical Shift**
   ```
   Thymus IP: 100, OP: 80, Paraspinous IP: 90, OP: 85
   ```

9. **Hepatic Steatosis**
   ```
   Liver IP: 100, OP: 80, Spleen IP: 90, OP: 88
   ```

10. **MRI Liver Iron Content**
    ```
    1.5T, R2*: 50 Hz
    ```

11. **Nodule Comparison & Sorting**
    ```
    4.1 x 2.9 x 3.3 cm (previously 3.9 x 2.8 x 3.0 cm)
    ```

12. **Number Statistics**
    ```
    3, 12, 5, 8, 15, 7
    ```

13. **Measurement Formatting**
    ```
    3.5 x 2.6 x 4.2 cm
    ```

14. **Fleischner 2017 Recommendations**
    ```
    Incidental right upper lobe solid noncalcified pulmonary nodule measuring 7 x 8 mm
    ```

15. **Contrast Premedication**
    - Interactive calculator for premedication timing

16. **NASCET Calculator**
    ```
    Distal ICA: 5mm, stenosis: 2mm
    ```

## Installation

### Compile from Source
1. Install AutoHotKey v1.1
2. Download source code
3. Compile using AHK Compiler
4. Transfer and run the generated .exe

## Usage

### Basic Usage
1. Highlight text in PowerScribe360 or NotePad
2. Right-click
3. Select calculation from menu
4. Review results in popup (automatically copied to clipboard)
5. To exit: Find green H icon in system tray → right-click → Exit

### References Management
1. **Adding References**
   - Right-click → References → Add Reference
   - Enter reference name and URL/file path
   - For files: Browse to select, file will be copied to References folder
   - For URLs: Enter complete web address

2. **Using References**
   - Right-click → References → [Select saved reference]
   - Files will open in default application
   - URLs will open in default browser
   - Most used references appear at top of menu

3. **Managing References**
   - Right-click → References → Remove References
   - Select references to remove
   - Confirm deletion

### AI Assistant Integration
1. **Using Built-in Prompts**
   - Highlight text (e.g., report section)
   - Right-click → AI Assistant → [Select prompt]
   - Choose to proceed with external AI warning
   - Promot and highlighted text opens in Claude or ChatGPT based on settings

2. **Managing Prompts**
   - Right-click → AI Assistant → Add Prompt
   - Enter prompt name and template. The promot name will be shown in the right click menu. The promot template will be appended to AI queries.
   - Use Manage Prompts to edit/remove
   - Default prompts include:
     - Edit Report (grammar, clarity, structure)
     - Generate Impression 

3. **AI Settings**
   - Right-click → AI Assistant → AI Settings
   - Select preferred AI (Claude/ChatGPT)
   - Restore default prompts if needed (removed any saved or custom prompts).

Note: The AI Assistant requires external access to AI services. Do not send PHI or sensitive information to external AI services. Attempts are made to sanatize but cannot he confirmed to work in every case. 

## Troubleshooting

- Check for duplicate instances (green H in system tray)
- Verify supported application (PowerScribe360/Notepad)
- Use pause feature if conflicting with other software
- Review input format against examples
- For Fleischner recommendations, ensure complete nodule description

## Contributing

We welcome contributions in:

- New calculation modules
- Text processing improvements
- RADS system implementations
- Language support
- PACS/EHR integration
- Automated data capture
- NLP capabilities
- UI enhancements
- Documentation
- Testing

To contribute:
1. Fork repository
2. Create feature branch
3. Commit changes
4. Submit pull request

## License

GPL - free for personal use, commercial use requires a license.

## Disclaimer

For educational and research purposes only. Not a substitute for professional medical judgment. Users must verify all calculations. The developers and software are not responsible for clinical decisions or errors. Do NOT send PHI to third parties.
