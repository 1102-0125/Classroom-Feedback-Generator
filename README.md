
# Classroom Feedback Form Generator

![GitHub repo size](https://img.shields.io/github/repo-size/1102-0125/classroom-feedback-generator)
![GitHub contributors](https://img.shields.io/github/contributors/1102-0125/classroom-feedback-generator)
![GitHub stars](https://img.shields.io/github/stars/1102-0125/classroom-feedback-generator?style=social)
![GitHub forks](https://img.shields.io/github/forks/1102-0125/classroom-feedback-generator?style=social)

A powerful desktop application for generating personalized student feedback forms from Excel data. Ideal for teachers and educators who need to create detailed, professional feedback reports efficiently.

![Application Screenshot](https://picsum.photos/800/600?random=1)

## Features

- **Excel Data Processing**: Reads student and course data from Excel spreadsheets
- **Customizable Column Mapping**: Configurable column settings for flexible data extraction
- **Parallel PDF Generation**: Uses multi-threading for fast batch processing
- **Personalized Reports**: Generates individualized feedback forms with student-specific data
- **Progress Tracking**: Real-time progress updates and detailed processing logs
- **HTML Template Support**: Fully customizable HTML templates for report formatting
- **Date Range Filtering**: Includes specified date ranges in generated reports

## Installation

### Prerequisites

- Python 3.7+
- Required Python packages (install via `pip`):
  - pandas
  - playwright
  - pypinyin
  - tkinter (usually included with Python)

### Steps

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/classroom-feedback-generator.git
   cd classroom-feedback-generator
   ```

2. Install dependencies:
   ```bash
   pip install pandas playwright pypinyin
   ```

3. Install Playwright browsers:
   ```bash
   playwright install
   ```

## Usage

1. Prepare your Excel file with student and course data
2. Run the application:
   ```bash
   python main.py
   ```
3. Configure settings:
   - Select your Excel input file
   - Choose output directory for generated PDFs
   - Set the date range for the feedback period
   - Adjust column mappings if needed
4. Click "Start Generation" to begin processing
5. Monitor progress and logs in the application window

## Configuration

The application uses an HTML template (`structure.html`) to format the generated PDFs. You can customize this template to match your school's branding or specific requirements. The template supports placeholders such as:

- `[name]`: Student's Chinese name
- `[engName]`: Student's English name
- `[pinyin]`: Pinyin transcription of the name
- `[date]`: Date range specified in the application
- `<div id="course-section"></div>`: Automatically replaced with course details

## Contributing

We welcome contributions from the community! To contribute:

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/your-feature`)
3. Commit your changes (`git commit -m 'Add some feature'`)
4. Push to the branch (`git push origin feature/your-feature`)
5. Open a Pull Request

### Development Guidelines

- Follow PEP 8 style guidelines for Python code
- Add unit tests for new features
- Update documentation as needed
- Ensure all tests pass before submitting a pull request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Thanks to all contributors who have helped improve this project
- Special thanks to the developers of pandas, playwright, and other dependencies
- Inspired by the needs of educators worldwide to streamline feedback processes

## Contact

If you have any questions or suggestions, please open an issue or contact the maintainers.
